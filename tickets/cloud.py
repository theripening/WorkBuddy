"""
WorkBuddyCloud API client.

All functions are no-ops when WORKBUDDY_CLOUD_URL is not set in settings.
Errors are logged but never raised — cloud sync is best-effort.
"""
import logging
from datetime import datetime, timezone

import requests
from django.conf import settings

logger = logging.getLogger(__name__)

_TIMEOUT = 5  # seconds


def _base_url():
    return getattr(settings, "WORKBUDDY_CLOUD_URL", None)


def _conv_id_for_ticket(ticket):
    """Return the conversation_id for a ticket's seed email, or first email."""
    email = ticket.emails.filter(is_seed=True).first() or ticket.emails.first()
    return email.conversation_id if email else None


# ---------------------------------------------------------------------------
# Outlook helpers
# ---------------------------------------------------------------------------

def get_my_email(namespace):
    """Read the current user's primary SMTP address from Outlook."""
    try:
        return namespace.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
    except Exception:
        try:
            return namespace.CurrentUser.Address
        except Exception:
            return None


def forward_to_assignee(ticket, assignee_email):
    """
    Forward seed email(s) to the assignee via Outlook COM, flagged for follow-up.
    Returns the sender's email address, or None on failure.
    """
    import pythoncom
    import win32com.client

    seed_ids = list(ticket.emails.filter(is_seed=True).values_list("outlook_id", flat=True))
    if not seed_ids:
        # Fall back to most recent email
        first = ticket.emails.first()
        if first:
            seed_ids = [first.outlook_id]
    if not seed_ids:
        logger.warning("Ticket %d: no emails to forward", ticket.pk)
        return None

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        my_email = get_my_email(namespace)

        for outlook_id in seed_ids:
            try:
                item = namespace.GetItemFromID(outlook_id)
                fwd = item.Forward()
                fwd.Recipients.Add(assignee_email)
                fwd.Recipients.ResolveAll()
                fwd.FlagStatus = 2  # olFlagMarked — creates a task flag on arrival
                fwd.Send()
                logger.info("Forwarded outlook_id=%s to %s", outlook_id, assignee_email)
            except Exception as e:
                logger.warning("Could not forward outlook_id=%s: %s", outlook_id, e)

        return my_email
    finally:
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# Cloud API calls
# ---------------------------------------------------------------------------

def push_ticket(ticket, assignee_email, my_email):
    """
    Create or update a shared ticket on WorkBuddyCloud.
    Called when a ticket is assigned.
    """
    url = _base_url()
    if not url:
        return
    conv_id = _conv_id_for_ticket(ticket)
    if not conv_id:
        logger.warning("Ticket %d: no conversation_id, cannot push to cloud", ticket.pk)
        return
    try:
        r = requests.post(
            f"{url}/api/tickets/",
            json={
                "conversation_id": conv_id,
                "subject": ticket.subject,
                "priority": ticket.priority or "medium",
                "assigner_email": my_email,
                "assignee_email": assignee_email,
            },
            timeout=_TIMEOUT,
        )
        r.raise_for_status()
        logger.info("Pushed ticket %d to cloud (conv=%s)", ticket.pk, conv_id[:16])
    except Exception as e:
        logger.warning("Cloud push failed for ticket %d: %s", ticket.pk, e)


def push_status(ticket, status, my_email):
    """Push a status update for a ticket to WorkBuddyCloud."""
    url = _base_url()
    if not url:
        return
    conv_id = _conv_id_for_ticket(ticket)
    if not conv_id:
        return
    try:
        r = requests.patch(
            f"{url}/api/tickets/{conv_id}/",
            json={"status": status},
            timeout=_TIMEOUT,
        )
        r.raise_for_status()
        logger.info("Pushed status=%s for ticket %d to cloud", status, ticket.pk)
    except Exception as e:
        logger.warning("Cloud status push failed for ticket %d: %s", ticket.pk, e)


def push_note(ticket, text, my_email):
    """Push a note for a ticket to WorkBuddyCloud."""
    url = _base_url()
    if not url:
        return
    conv_id = _conv_id_for_ticket(ticket)
    if not conv_id:
        return
    try:
        r = requests.post(
            f"{url}/api/tickets/{conv_id}/notes/",
            json={"author_email": my_email, "text": text},
            timeout=_TIMEOUT,
        )
        r.raise_for_status()
        logger.info("Pushed note to cloud for ticket %d", ticket.pk)
    except Exception as e:
        logger.warning("Cloud note push failed for ticket %d: %s", ticket.pk, e)


def pull_cloud_notes(my_email):
    """
    Pull all cloud tickets involving me (assigned to me or assigned by me).
    Returns a list of cloud ticket dicts with notes.
    """
    url = _base_url()
    if not url:
        return []
    results = []
    for param in ("email", "assigner"):
        try:
            r = requests.get(
                f"{url}/api/tickets/",
                params={param: my_email},
                timeout=_TIMEOUT,
            )
            r.raise_for_status()
            results.extend(r.json().get("tickets", []))
        except Exception as e:
            logger.warning("Cloud pull failed (%s=%s): %s", param, my_email, e)
    # Deduplicate by conversation_id
    seen = {}
    for t in results:
        seen[t["conversation_id"]] = t
    return list(seen.values())


def push_todo(ticket, todo):
    """Push a todo item to WorkBuddyCloud."""
    url = _base_url()
    if not url:
        return
    conv_id = _conv_id_for_ticket(ticket)
    if not conv_id:
        return
    try:
        r = requests.post(
            f"{url}/api/tickets/{conv_id}/todos/",
            json={"external_id": todo.pk, "text": todo.what},
            timeout=_TIMEOUT,
        )
        r.raise_for_status()
        logger.info("Pushed todo %d to cloud (ticket %d)", todo.pk, ticket.pk)
    except Exception as e:
        logger.warning("Cloud push_todo failed for todo %d: %s", todo.pk, e)


def complete_todo(todo):
    """Mark a cloud todo as completed."""
    url = _base_url()
    if not url:
        return
    try:
        r = requests.post(
            f"{url}/api/todos/{todo.pk}/complete/",
            json={"completed_by": "assigner"},
            timeout=_TIMEOUT,
        )
        r.raise_for_status()
        logger.info("Completed cloud todo %d", todo.pk)
    except Exception as e:
        logger.warning("Cloud complete_todo failed for todo %d: %s", todo.pk, e)


def push_waiting(ticket, waiting):
    """Push a waiting-on item to WorkBuddyCloud."""
    url = _base_url()
    if not url:
        return
    conv_id = _conv_id_for_ticket(ticket)
    if not conv_id:
        return
    try:
        r = requests.post(
            f"{url}/api/tickets/{conv_id}/waiting/",
            json={
                "external_id": waiting.pk,
                "what": waiting.what,
                "from_who": waiting.from_who,
                "expected_date": waiting.expected_date.isoformat() if waiting.expected_date else None,
            },
            timeout=_TIMEOUT,
        )
        r.raise_for_status()
        logger.info("Pushed waiting %d to cloud (ticket %d)", waiting.pk, ticket.pk)
    except Exception as e:
        logger.warning("Cloud push_waiting failed for waiting %d: %s", waiting.pk, e)


def resolve_waiting(waiting):
    """Mark a cloud waiting-on as resolved."""
    url = _base_url()
    if not url:
        return
    try:
        r = requests.post(
            f"{url}/api/waiting/{waiting.pk}/resolve/",
            json={},
            timeout=_TIMEOUT,
        )
        r.raise_for_status()
        logger.info("Resolved cloud waiting %d", waiting.pk)
    except Exception as e:
        logger.warning("Cloud resolve_waiting failed for waiting %d: %s", waiting.pk, e)


def sync_cloud_notes(my_email):
    """
    Pull cloud tickets and store any new notes locally as CloudNote records.
    Returns count of new notes saved.
    """
    from .models import Ticket, TicketEmail, CloudNote

    cloud_tickets = pull_cloud_notes(my_email)
    if not cloud_tickets:
        return 0

    # Build a map of conversation_id -> local Ticket
    conv_ids = [t["conversation_id"] for t in cloud_tickets]
    email_qs = TicketEmail.objects.filter(conversation_id__in=conv_ids).values("conversation_id", "ticket_id")
    conv_to_ticket = {row["conversation_id"]: row["ticket_id"] for row in email_qs}

    saved = 0
    for cloud_ticket in cloud_tickets:
        ticket_id = conv_to_ticket.get(cloud_ticket["conversation_id"])
        if not ticket_id:
            continue
        for note in cloud_ticket.get("notes", []):
            created_at = datetime.fromisoformat(note["created_at"])
            if created_at.tzinfo is None:
                created_at = created_at.replace(tzinfo=timezone.utc)
            _, created = CloudNote.objects.get_or_create(
                ticket_id=ticket_id,
                author_email=note["author_email"],
                created_at=created_at,
                defaults={"text": note["text"]},
            )
            if created:
                saved += 1

    if saved:
        logger.info("Cloud sync: %d new note(s) pulled", saved)
    return saved
