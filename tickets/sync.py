from datetime import datetime
from django.conf import settings
from .models import Ticket, TicketEmail

TRACKED_FOLDER_NAME = "Tracked"
OL_MAIL_ITEM = 43  # Outlook MailItem class constant
OL_FLAG_MARKED = 2  # Outlook olFlagMarked constant
PR_BODY_PREVIEW = "http://schemas.microsoft.com/mapi/proptag/0x0071001E"


def _parse_received_time(received_raw):
    return datetime(
        received_raw.year,
        received_raw.month,
        received_raw.day,
        received_raw.hour,
        received_raw.minute,
        received_raw.second,
        tzinfo=received_raw.tzinfo,
    )


def _safe_sender(item):
    try:
        return item.SenderName or item.SenderEmailAddress or ""
    except Exception:
        return ""


def _safe_body_preview(item):
    # PR_BODY_PREVIEW is pre-cached on the server — much faster than .Body
    try:
        preview = item.PropertyAccessor.GetProperty(PR_BODY_PREVIEW)
        if preview:
            return str(preview)[:300].strip()
    except Exception:
        pass
    # Fallback to .Body if MAPI property not available
    try:
        return (item.Body or "")[:300].strip()
    except Exception:
        return ""


def _build_email_obj(ticket, mail_item):
    """Build an unsaved TicketEmail instance."""
    return TicketEmail(
        ticket=ticket,
        outlook_id=mail_item.EntryID,
        conversation_id=mail_item.ConversationID,
        subject=mail_item.Subject or "(no subject)",
        sender=_safe_sender(mail_item),
        received_at=_parse_received_time(mail_item.ReceivedTime),
        body_preview=_safe_body_preview(mail_item),
    )


def _sync_conversation(namespace, seed_item, ticket, known_outlook_ids):
    """
    Walk all items in the conversation and collect unsaved TicketEmail
    objects for any EntryIDs not already in known_outlook_ids.
    Updates known_outlook_ids in place.
    Returns list of unsaved TicketEmail instances.
    """
    new_emails = []
    try:
        conversation = seed_item.GetConversation()
        if conversation is None:
            entry_id = seed_item.EntryID
            if entry_id not in known_outlook_ids:
                new_emails.append(_build_email_obj(ticket, seed_item))
                known_outlook_ids.add(entry_id)
            return new_emails

        # Collect all EntryIDs in one pass — no GetItemFromID yet
        table = conversation.GetTable()
        table.Columns.RemoveAll()
        table.Columns.Add("EntryID")

        entry_ids = []
        while not table.EndOfTable:
            try:
                entry_ids.append(table.GetNextRow()["EntryID"])
            except Exception:
                continue

        # Only fetch COM objects for entries we don't already have
        for entry_id in entry_ids:
            if entry_id in known_outlook_ids:
                continue
            try:
                mail_item = namespace.GetItemFromID(entry_id)
                if mail_item.Class != OL_MAIL_ITEM:
                    continue
                new_emails.append(_build_email_obj(ticket, mail_item))
                known_outlook_ids.add(entry_id)
            except Exception:
                continue

    except Exception:
        # GetConversation not available — fall back to seed only
        entry_id = seed_item.EntryID
        if entry_id not in known_outlook_ids:
            new_emails.append(_build_email_obj(ticket, seed_item))
            known_outlook_ids.add(entry_id)

    return new_emails


def sync_tracked_folder():
    """
    Connect to the running Outlook instance, find the 'Tracked' subfolder
    under Inbox, and sync emails + their full conversations into the DB.
    Returns (new_tickets, new_emails) counts.
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

        track_mode = getattr(settings, "TRACK_MODE", "folder")

        # Pre-load all known data to avoid per-item DB queries
        known_convs = {}
        for conv_id, ticket_id in TicketEmail.objects.values_list("conversation_id", "ticket_id"):
            known_convs[conv_id] = ticket_id

        known_outlook_ids = set(TicketEmail.objects.values_list("outlook_id", flat=True))

        tickets_cache = {}

        # Collect seed items from configured sources, deduplicating by ConversationID
        seen_conv_ids = set()
        convs_to_process = []  # list of (conv_id, entry_id, subject)

        def _collect_from_items(mail_items):
            count = mail_items.Count
            for i in range(1, count + 1):
                try:
                    item = mail_items.Item(i)
                    conv_id = item.ConversationID
                    if conv_id not in seen_conv_ids:
                        seen_conv_ids.add(conv_id)
                        convs_to_process.append((conv_id, item.EntryID, item.Subject or "(no subject)"))
                except Exception:
                    continue

        # --- Tracked subfolder ---
        if track_mode in ("folder", "both"):
            tracked = None
            for folder in inbox.Folders:
                if folder.Name == TRACKED_FOLDER_NAME:
                    tracked = folder
                    break
            if tracked is None and track_mode == "folder":
                raise ValueError(
                    f'Outlook folder "{TRACKED_FOLDER_NAME}" not found under Inbox. '
                    "Create it and move emails there, or set TRACK_MODE = 'flag'."
                )
            if tracked is not None:
                _collect_from_items(tracked.Items.Restrict("[MessageClass] = 'IPM.Note'"))

        # --- Flagged emails in Inbox ---
        if track_mode in ("flag", "both"):
            flagged = inbox.Items.Restrict(
                f"[FlagStatus] = {OL_FLAG_MARKED} AND [MessageClass] = 'IPM.Note'"
            )
            _collect_from_items(flagged)

        new_tickets = 0
        to_bulk_create = []

        for conv_id, entry_id, subject in convs_to_process:
            try:
                if conv_id in known_convs:
                    ticket_id = known_convs[conv_id]
                    if ticket_id not in tickets_cache:
                        tickets_cache[ticket_id] = Ticket.objects.get(pk=ticket_id)
                    ticket = tickets_cache[ticket_id]
                else:
                    ticket = Ticket.objects.create(subject=subject)
                    known_convs[conv_id] = ticket.pk
                    tickets_cache[ticket.pk] = ticket
                    new_tickets += 1

                # Re-fetch the seed item by EntryID to avoid COM pointer reuse issues
                seed_item = namespace.GetItemFromID(entry_id)
                new_email_objs = _sync_conversation(namespace, seed_item, ticket, known_outlook_ids)
                to_bulk_create.extend(new_email_objs)

            except Exception:
                continue

        # Bulk insert all new emails in one DB round-trip
        TicketEmail.objects.bulk_create(to_bulk_create, ignore_conflicts=True)

        return new_tickets, len(to_bulk_create)

    finally:
        pythoncom.CoUninitialize()
