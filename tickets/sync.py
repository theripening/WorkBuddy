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


def _upsert_email(ticket, mail_item):
    """
    Create a TicketEmail for mail_item if it doesn't exist yet.
    Returns True if a new record was created.
    """
    outlook_id = mail_item.EntryID
    _, created = TicketEmail.objects.get_or_create(
        outlook_id=outlook_id,
        defaults={
            "ticket": ticket,
            "conversation_id": mail_item.ConversationID,
            "subject": mail_item.Subject or "(no subject)",
            "sender": _safe_sender(mail_item),
            "received_at": _parse_received_time(mail_item.ReceivedTime),
            "body_preview": _safe_body_preview(mail_item),
        },
    )
    return created


def _sync_conversation(namespace, seed_item, ticket):
    """
    Walk all items in the same conversation as seed_item and upsert them.
    Returns count of newly created TicketEmails.
    """
    new_count = 0
    try:
        conversation = seed_item.GetConversation()
        if conversation is None:
            # Fallback: just upsert the seed item itself
            if _upsert_email(ticket, seed_item):
                new_count += 1
            return new_count

        table = conversation.GetTable()
        table.Columns.RemoveAll()
        table.Columns.Add("EntryID")

        while not table.EndOfTable:
            row = table.GetNextRow()
            try:
                entry_id = row["EntryID"]
                mail_item = namespace.GetItemFromID(entry_id)
                if mail_item.Class != OL_MAIL_ITEM:
                    continue
                if _upsert_email(ticket, mail_item):
                    new_count += 1
            except Exception:
                continue

    except Exception:
        # GetConversation not available — upsert seed only
        if _upsert_email(ticket, seed_item):
            new_count += 1

    return new_count


def sync_tracked_folder():
    """
    Connect to the running Outlook instance, find the 'Tracked' subfolder
    under Inbox, and sync emails + their full conversations into the DB.
    Returns (new_tickets, new_emails) counts.
    """
    import win32com.client

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

    new_tickets = 0
    new_emails = 0

    for item in tracked.Items:
        try:
            if item.Class != OL_MAIL_ITEM:
                continue

            conv_id = item.ConversationID

            # Find or create the ticket for this conversation
            existing = TicketEmail.objects.filter(conversation_id=conv_id).first()
            if existing:
                ticket = existing.ticket
                created_ticket = False
            else:
                ticket = Ticket.objects.create(
                    subject=item.Subject or "(no subject)",
                )
                created_ticket = True

            if created_ticket:
                new_tickets += 1

            new_emails += _sync_conversation(namespace, item, ticket)

        except Exception:
            continue

    return new_tickets, new_emails
