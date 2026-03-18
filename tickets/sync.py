from datetime import datetime
from .models import Ticket, TicketEmail

TRACKED_FOLDER_NAME = "Tracked"
OL_MAIL_ITEM = 43  # Outlook MailItem class constant


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
    try:
        return (item.Body or "")[:300].strip()
    except Exception:
        return ""


def _create_email(ticket, mail_item):
    TicketEmail.objects.create(
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
    Walk all items in the conversation and create TicketEmails for any
    EntryIDs not already in known_outlook_ids (mutated in place).
    Returns count of newly created TicketEmails.
    """
    new_count = 0
    try:
        conversation = seed_item.GetConversation()
        if conversation is None:
            if seed_item.EntryID not in known_outlook_ids:
                _create_email(ticket, seed_item)
                known_outlook_ids.add(seed_item.EntryID)
                new_count += 1
            return new_count

        # Collect all EntryIDs from the conversation table in one pass
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
                _create_email(ticket, mail_item)
                known_outlook_ids.add(entry_id)
                new_count += 1
            except Exception:
                continue

    except Exception:
        # GetConversation not available — upsert seed only
        if seed_item.EntryID not in known_outlook_ids:
            _create_email(ticket, seed_item)
            known_outlook_ids.add(seed_item.EntryID)
            new_count += 1

    return new_count


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

        tracked = None
        for folder in inbox.Folders:
            if folder.Name == TRACKED_FOLDER_NAME:
                tracked = folder
                break

        if tracked is None:
            raise ValueError(
                f'Outlook folder "{TRACKED_FOLDER_NAME}" not found under Inbox. '
                "Create it and move emails there before syncing."
            )

        # Pre-load all known data to avoid per-item DB queries
        known_convs = {}
        for conv_id, ticket_id in TicketEmail.objects.values_list("conversation_id", "ticket_id"):
            known_convs[conv_id] = ticket_id

        # Set of all known outlook_ids — checked before any GetItemFromID call
        known_outlook_ids = set(TicketEmail.objects.values_list("outlook_id", flat=True))

        # Cache of ticket_id -> Ticket to avoid repeated DB fetches
        tickets_cache = {}

        # Deduplicate: only process each conversation once per sync run
        seen_conv_ids = set()
        convs_to_process = []  # list of (conv_id, seed_item)

        for item in tracked.Items:
            try:
                if item.Class != OL_MAIL_ITEM:
                    continue
                conv_id = item.ConversationID
                if conv_id not in seen_conv_ids:
                    seen_conv_ids.add(conv_id)
                    convs_to_process.append((conv_id, item))
            except Exception:
                continue

        new_tickets = 0
        new_emails = 0

        for conv_id, seed_item in convs_to_process:
            try:
                if conv_id in known_convs:
                    ticket_id = known_convs[conv_id]
                    if ticket_id not in tickets_cache:
                        tickets_cache[ticket_id] = Ticket.objects.get(pk=ticket_id)
                    ticket = tickets_cache[ticket_id]
                else:
                    ticket = Ticket.objects.create(subject=seed_item.Subject or "(no subject)")
                    known_convs[conv_id] = ticket.pk
                    tickets_cache[ticket.pk] = ticket
                    new_tickets += 1

                new_emails += _sync_conversation(namespace, seed_item, ticket, known_outlook_ids)

            except Exception:
                continue

        return new_tickets, new_emails

    finally:
        pythoncom.CoUninitialize()
