import time
from datetime import datetime
from django.conf import settings
from .models import Ticket, TicketEmail

TRACKED_FOLDER_NAME = "Tracked"
OL_MAIL_ITEM = 43  # Outlook MailItem class constant
OL_FLAG_MARKED = 2  # Outlook olFlagMarked constant
PR_BODY_PREVIEW = "http://schemas.microsoft.com/mapi/proptag/0x0071001E"

# All columns fetched in a single GetTable call — eliminates GetItemFromID per email
_TABLE_COLUMNS = ["EntryID", "ConversationID", "MessageClass", "Subject", "SenderName", "ReceivedTime", "SentOn"]


def _t(label, start):
    print(f"  [{label}] {time.perf_counter() - start:.3f}s")


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
    """Build TicketEmail from a live COM mail item (used in fallback paths only)."""
    return TicketEmail(
        ticket=ticket,
        outlook_id=mail_item.EntryID,
        conversation_id=mail_item.ConversationID,
        subject=mail_item.Subject or "(no subject)",
        sender=_safe_sender(mail_item),
        received_at=_parse_received_time(mail_item.ReceivedTime),
        body_preview=_safe_body_preview(mail_item),
    )


def _build_email_from_row(ticket, row):
    """Build TicketEmail directly from a conversation table row — zero extra COM calls."""
    body = ""
    try:
        raw = row[PR_BODY_PREVIEW]
        if raw:
            body = str(raw)[:300].strip()
    except Exception:
        pass
    # ReceivedTime may be None for sent/draft items — fall back to SentOn
    received_raw = row.get("ReceivedTime") or row.get("SentOn")
    if received_raw is None:
        raise ValueError(f"No ReceivedTime or SentOn for EntryID {row.get('EntryID')}")
    return TicketEmail(
        ticket=ticket,
        outlook_id=row["EntryID"],
        conversation_id=row.get("ConversationID", ""),
        subject=row.get("Subject") or "(no subject)",
        sender=row.get("SenderName") or "",
        received_at=_parse_received_time(received_raw),
        body_preview=body,
    )


def _sync_conversation(namespace, seed_entry_id, ticket, known_outlook_ids):
    """
    Walk all items in the conversation and collect unsaved TicketEmail objects.

    Fetches all needed columns in a single GetTable call so no GetItemFromID
    is required per email — only one GetItemFromID for the seed (needed to call
    GetConversation), then everything else comes from table rows.

    Updates known_outlook_ids in place.
    Returns (list of unsaved TicketEmail instances, dict of timing info).
    """
    new_emails = []
    timings = {}

    t0 = time.perf_counter()
    try:
        seed_item = namespace.GetItemFromID(seed_entry_id)
    except Exception:
        return new_emails, {"error": "GetItemFromID failed"}
    timings["GetItemFromID"] = time.perf_counter() - t0

    try:
        t0 = time.perf_counter()
        conversation = seed_item.GetConversation()
        timings["GetConversation"] = time.perf_counter() - t0

        if conversation is None:
            timings["conv_none"] = True
            if seed_entry_id not in known_outlook_ids:
                new_emails.append(_build_email_obj(ticket, seed_item))
                known_outlook_ids.add(seed_entry_id)
            return new_emails, timings

        t0 = time.perf_counter()
        table = conversation.GetTable()
        table.Columns.RemoveAll()
        for col in _TABLE_COLUMNS:
            table.Columns.Add(col)
        table.Columns.Add(PR_BODY_PREVIEW)
        timings["GetTable+Columns"] = time.perf_counter() - t0

        t0 = time.perf_counter()
        rows = []
        while not table.EndOfTable:
            try:
                rows.append(table.GetNextRow())
            except Exception:
                continue
        timings["TableWalk"] = time.perf_counter() - t0
        timings["row_count"] = len(rows)

        # Early-out: all entry IDs already known — nothing to do
        if all(row["EntryID"] in known_outlook_ids for row in rows):
            timings["early_out"] = True
            return [], timings

        skipped_class = 0
        for row in rows:
            entry_id = row["EntryID"]
            if entry_id in known_outlook_ids:
                continue
            if row.get("MessageClass", "") != "IPM.Note":
                skipped_class += 1
                continue
            try:
                new_emails.append(_build_email_from_row(ticket, row))
                known_outlook_ids.add(entry_id)
            except Exception as e:
                print(f"    SKIP row {entry_id[:20]}... — {e}")
                continue
        if skipped_class:
            timings["skipped_non_mail"] = skipped_class

    except Exception:
        # GetConversation not available — fall back to seed item only
        if seed_entry_id not in known_outlook_ids:
            try:
                new_emails.append(_build_email_obj(ticket, seed_item))
                known_outlook_ids.add(seed_entry_id)
            except Exception:
                pass

    return new_emails, timings


def sync_tracked_folder():
    """
    Connect to the running Outlook instance and sync tracked emails + their full
    conversations into the DB. Returns (new_tickets, new_emails) counts.

    Tracking sources are controlled by TRACK_MODE in settings:
      "folder" — Inbox > Tracked subfolder
      "flag"   — flagged (red flag) emails in Inbox
      "both"   — either source (default)
    """
    import pythoncom
    import win32com.client

    sync_start = time.perf_counter()
    print("\n=== WorkBuddy Sync Start ===")

    pythoncom.CoInitialize()
    try:
        t0 = time.perf_counter()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        _t("Outlook connect", t0)

        track_mode = getattr(settings, "TRACK_MODE", "folder")
        print(f"  TRACK_MODE = {track_mode!r}")

        # Pre-load all known data in two DB queries
        t0 = time.perf_counter()
        known_convs = {}
        for conv_id, ticket_id in TicketEmail.objects.values_list("conversation_id", "ticket_id"):
            known_convs[conv_id] = ticket_id
        known_outlook_ids = set(TicketEmail.objects.values_list("outlook_id", flat=True))
        _t(f"DB pre-load ({len(known_convs)} convs, {len(known_outlook_ids)} emails)", t0)

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
            t0 = time.perf_counter()
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
            _t(f"Folder collection ({len(convs_to_process)} convs so far)", t0)

        # --- Flagged emails via To-Do special folder ---
        # olFolderToDo aggregates flagged items from all folders — no recursion needed.
        # GetTable() and FlagStatus in Restrict both fail on this virtual folder, so
        # we Restrict by MessageClass only (works), then use ConversationID presence
        # to identify email items (tasks have no ConversationID).
        if track_mode in ("flag", "both"):
            t0 = time.perf_counter()
            before = len(convs_to_process)
            try:
                todo_folder = namespace.GetDefaultFolder(28)  # 28 = olFolderToDo
                mail_todos = todo_folder.Items.Restrict("[MessageClass] = 'IPM.Note'")
                count = mail_todos.Count
                print(f"  To-Do folder: {count} mail item(s)")
                for i in range(1, count + 1):
                    try:
                        item = mail_todos.Item(i)
                        if item.FlagStatus != OL_FLAG_MARKED:  # skip completed flags
                            continue
                        conv_id = item.ConversationID
                        if conv_id and conv_id not in seen_conv_ids:
                            seen_conv_ids.add(conv_id)
                            convs_to_process.append((conv_id, item.EntryID, item.Subject or "(no subject)"))
                    except Exception:
                        continue
            except Exception as e:
                print(f"  To-Do folder error: {e}")
            _t(f"Flag collection (+{len(convs_to_process) - before} convs, {len(convs_to_process)} total)", t0)

        print(f"  Conversations to process: {len(convs_to_process)}")

        new_tickets = 0
        to_bulk_create = []

        for idx, (conv_id, entry_id, subject) in enumerate(convs_to_process, 1):
            conv_start = time.perf_counter()
            try:
                is_new = conv_id not in known_convs
                if not is_new:
                    ticket_id = known_convs[conv_id]
                    if ticket_id not in tickets_cache:
                        tickets_cache[ticket_id] = Ticket.objects.get(pk=ticket_id)
                    ticket = tickets_cache[ticket_id]
                else:
                    ticket = Ticket.objects.create(subject=subject)
                    known_convs[conv_id] = ticket.pk
                    tickets_cache[ticket.pk] = ticket
                    new_tickets += 1

                new_email_objs, timings = _sync_conversation(namespace, entry_id, ticket, known_outlook_ids)
                to_bulk_create.extend(new_email_objs)

                elapsed = time.perf_counter() - conv_start
                early = timings.get("early_out", False)
                conv_none = timings.get("conv_none", False)
                rows = timings.get("row_count", "?")
                if conv_none:
                    label = "NONE"  # GetConversation() returned None — only seed stored
                elif is_new:
                    label = "NEW "
                elif early:
                    label = "skip"
                else:
                    label = "upd "
                print(
                    f"  [{idx:3d}/{len(convs_to_process)}] {label} | "
                    f"{elapsed:.3f}s total | "
                    f"GetItem={timings.get('GetItemFromID', 0):.3f}s "
                    f"GetConv={timings.get('GetConversation', 0):.3f}s "
                    f"Table={timings.get('GetTable+Columns', 0):.3f}s "
                    f"Walk={timings.get('TableWalk', 0):.3f}s "
                    f"rows={rows} new={len(new_email_objs)} | "
                    f"{subject[:50]}"
                )

            except Exception as e:
                print(f"  [{idx:3d}/{len(convs_to_process)}] ERROR: {e} | {subject[:50]}")
                continue

        t0 = time.perf_counter()
        TicketEmail.objects.bulk_create(to_bulk_create, ignore_conflicts=True)
        _t(f"bulk_create ({len(to_bulk_create)} emails)", t0)

        total = time.perf_counter() - sync_start
        print(f"=== Sync done in {total:.2f}s — {new_tickets} new tickets, {len(to_bulk_create)} new emails ===\n")
        return new_tickets, len(to_bulk_create)

    finally:
        pythoncom.CoUninitialize()
