import time
from datetime import datetime
from django.conf import settings
from .models import Ticket, TicketEmail

TRACKED_FOLDER_NAME = "Tracked"
OL_FLAG_MARKED = 2  # Outlook olFlagMarked constant
# Columns fetched in a single GetTable call — no body preview here,
# we fetch .Body via GetItemFromID when needed instead of MAPI proptags
# which return unpredictable binary types across Outlook versions.
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
    try:
        return (item.Body or "")[:300].strip()
    except Exception:
        return ""


def _build_email_obj(ticket, mail_item):
    """Build TicketEmail from a live COM mail item."""
    return TicketEmail(
        ticket=ticket,
        outlook_id=mail_item.EntryID,
        conversation_id=mail_item.ConversationID,
        subject=mail_item.Subject or "(no subject)",
        sender=_safe_sender(mail_item),
        received_at=_parse_received_time(mail_item.ReceivedTime),
        body_preview=_safe_body_preview(mail_item),
    )


def _row_get(row, key, default=None):
    """Safe key access on an Outlook COM Row object — Row supports [] but not .get()."""
    try:
        return row[key]
    except Exception:
        return default


def _build_email_from_row(ticket, row, namespace=None, conv_id="", fetch_body=True):
    """Build TicketEmail from a conversation table row."""
    body = ""
    if fetch_body and namespace is not None:
        try:
            live = namespace.GetItemFromID(row["EntryID"])
            body = (live.Body or "")[:300].strip()
        except Exception:
            pass
    received_raw = _row_get(row, "ReceivedTime") or _row_get(row, "SentOn")
    if received_raw is None:
        raise ValueError(f"No ReceivedTime or SentOn for EntryID {_row_get(row, 'EntryID')}")
    return TicketEmail(
        ticket=ticket,
        outlook_id=row["EntryID"],
        conversation_id=_row_get(row, "ConversationID") or conv_id,
        subject=_row_get(row, "Subject") or "(no subject)",
        sender=_row_get(row, "SenderName") or "",
        received_at=_parse_received_time(received_raw),
        body_preview=body,
    )


def _collect_conversation_emails(namespace, seed_entry_id, ticket, sent_by_conv, known_outlook_ids=None):
    """
    Return only NEW TicketEmail objects for the conversation containing seed_entry_id.

    Walks the full conversation table every sync to discover new emails, but skips
    GetItemFromID (body fetch) for any EntryID already in known_outlook_ids.

    Two sources:
      1. GetConversation().GetTable() — Outlook's built-in conversation view
      2. sent_by_conv dict (pre-loaded once) — catches sent replies that
         GetTable() misses when they live in Sent Items

    Returns (list of new TicketEmail instances, dict of timing info).
    """
    if known_outlook_ids is None:
        known_outlook_ids = set()
    timings = {}
    emails_by_entry_id = {}  # EntryID -> TicketEmail, deduped in memory

    print(f"    GetItemFromID...", end=" ", flush=True)
    t0 = time.perf_counter()
    try:
        seed_item = namespace.GetItemFromID(seed_entry_id)
    except Exception:
        print("FAILED")
        return [], {"error": "GetItemFromID failed"}
    timings["GetItemFromID"] = time.perf_counter() - t0
    print(f"{timings['GetItemFromID']:.3f}s")

    conv_id = seed_item.ConversationID

    # --- Source 1: GetConversation().GetTable() ---
    print(f"    GetConversation...", end=" ", flush=True)
    try:
        t0 = time.perf_counter()
        conversation = seed_item.GetConversation()
        timings["GetConversation"] = time.perf_counter() - t0
        print(f"{timings['GetConversation']:.3f}s", end=" ")

        if conversation is not None:
            t0 = time.perf_counter()
            table = conversation.GetTable()
            table.Columns.RemoveAll()
            for col in _TABLE_COLUMNS:
                table.Columns.Add(col)
            timings["GetTable"] = time.perf_counter() - t0

            t0 = time.perf_counter()
            rows = []
            while not table.EndOfTable:
                try:
                    rows.append(table.GetNextRow())
                except Exception:
                    continue
            timings["TableWalk"] = time.perf_counter() - t0
            timings["table_rows"] = len(rows)
            print(f"table_rows={len(rows)}")

            skipped_class = {}
            skipped_known = 0
            for row in rows:
                msg_class = _row_get(row, "MessageClass") or ""
                if msg_class != "IPM.Note":
                    skipped_class[msg_class] = skipped_class.get(msg_class, 0) + 1
                    continue
                try:
                    entry_id = row["EntryID"]
                    if entry_id in emails_by_entry_id:
                        continue
                    if entry_id in known_outlook_ids:
                        skipped_known += 1
                        continue
                    emails_by_entry_id[entry_id] = _build_email_from_row(ticket, row, namespace, conv_id)
                except Exception as e:
                    print(f"    SKIP table row: {e}")
            timings["skipped_class"] = skipped_class
            timings["skipped_known"] = skipped_known
        else:
            print("None")
            timings["conv_none"] = True

    except Exception as e:
        print(f"ERROR: {e}")
        timings["conv_err"] = str(e)

    # --- Source 2: pre-loaded Sent Items dict lookup (O(1) per conversation) ---
    sent_rows = sent_by_conv.get(conv_id, [])
    sent_matches = 0
    sent_skipped_known = 0
    for row_dict in sent_rows:
        entry_id = row_dict.get("EntryID")
        if not entry_id or entry_id in emails_by_entry_id:
            continue
        if entry_id in known_outlook_ids:
            sent_skipped_known += 1
            continue
        try:
            emails_by_entry_id[entry_id] = _build_email_from_row(ticket, row_dict, namespace, conv_id)
            sent_matches += 1
        except Exception as e:
            print(f"    SKIP sent row: {e}")
    timings["sent_matches"] = sent_matches
    timings["sent_skipped_known"] = sent_skipped_known
    print(f"    Sent lookup: {sent_matches} new, {sent_skipped_known} already known, from {len(sent_rows)} in conv")

    # --- Last resort: at minimum store the seed itself ---
    if not emails_by_entry_id and seed_entry_id not in emails_by_entry_id:
        try:
            emails_by_entry_id[seed_entry_id] = _build_email_obj(ticket, seed_item)
            timings["seed_fallback"] = True
        except Exception:
            pass

    return list(emails_by_entry_id.values()), timings


def sync_tracked_folder():
    """
    Connect to the running Outlook instance and sync tracked emails + their full
    conversations into the DB. Returns (new_tickets, new_emails) counts.

    Tracking sources are controlled by TRACK_MODE in settings:
      "folder" — Inbox > Tracked subfolder
      "flag"   — flagged (red flag) emails in Inbox
      "both"   — either source (default)

    Each sync fetches the full conversation fresh — no incremental tracking.
    bulk_create(ignore_conflicts=True) deduplicates at the DB level.
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

        # Pre-load known conversations and email IDs from DB
        t0 = time.perf_counter()
        known_convs = {}
        known_outlook_ids = set()
        for outlook_id, conv_id, ticket_id in TicketEmail.objects.values_list("outlook_id", "conversation_id", "ticket_id"):
            known_convs[conv_id] = ticket_id
            known_outlook_ids.add(outlook_id)
        _t(f"DB pre-load ({len(known_convs)} known conversations, {len(known_outlook_ids)} known emails)", t0)

        # Pre-load Sent Items once into {conv_id: [row_dict, ...]} — avoids
        # scanning 5k+ items per conversation during the main loop.
        t0 = time.perf_counter()
        sent_by_conv = {}
        try:
            sent_folder = namespace.GetDefaultFolder(5)  # olFolderSentMail
            sent_table = sent_folder.GetTable()
            sent_table.Columns.RemoveAll()
            for col in _TABLE_COLUMNS:
                sent_table.Columns.Add(col)
            scanned = 0
            no_class = 0
            no_conv = 0
            while not sent_table.EndOfTable:
                try:
                    row = sent_table.GetNextRow()
                    scanned += 1
                    msg_class = _row_get(row, "MessageClass") or ""
                    if msg_class != "IPM.Note":
                        no_class += 1
                        continue
                    cid = _row_get(row, "ConversationID")
                    if not cid:
                        no_conv += 1
                        continue
                    row_dict = {col: _row_get(row, col) for col in _TABLE_COLUMNS}
                    sent_by_conv.setdefault(cid, []).append(row_dict)
                except Exception as e:
                    print(f"  Sent row error: {e}")
                    continue
            print(f"  Sent pre-load: scanned={scanned} skipped_class={no_class} no_conv_id={no_conv}")
        except Exception as e:
            print(f"  Sent Items pre-load error: {e}")
        sent_total = sum(len(v) for v in sent_by_conv.values())
        _t(f"Sent Items pre-load ({sent_total} emails across {len(sent_by_conv)} convs)", t0)

        tickets_cache = {}

        # Collect seed items from configured sources, deduplicating by ConversationID
        seen_conv_ids = set()
        convs_to_process = []  # list of (conv_id, entry_id, subject)

        def _collect_from_items(mail_items):
            count = mail_items.Count
            for i in range(1, count + 1):
                try:
                    item = mail_items.Item(i)
                    cid = item.ConversationID
                    if cid not in seen_conv_ids:
                        seen_conv_ids.add(cid)
                        convs_to_process.append((cid, item.EntryID, item.Subject or "(no subject)"))
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
        # we Restrict by MessageClass only, then check FlagStatus in Python.
        # Re-fetch each seed via GetItemFromID to get a real store-bound item so
        # GetConversation() works correctly.
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
                        if item.FlagStatus != OL_FLAG_MARKED:
                            continue
                        cid = item.ConversationID
                        if cid and cid not in seen_conv_ids:
                            seen_conv_ids.add(cid)
                            # Re-fetch from MAPI — virtual folder items may have
                            # shortcut EntryIDs that cause GetConversation() to fail
                            real_entry_id = namespace.GetItemFromID(item.EntryID).EntryID
                            convs_to_process.append((cid, real_entry_id, item.Subject or "(no subject)"))
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
            print(f"  [{idx:3d}/{len(convs_to_process)}] {subject[:60]}")
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

                email_objs, timings = _collect_conversation_emails(namespace, entry_id, ticket, sent_by_conv, known_outlook_ids)
                to_bulk_create.extend(email_objs)

                elapsed = time.perf_counter() - conv_start
                label = "NEW " if is_new else "sync"
                table_rows = timings.get("table_rows", "?")
                sent_matches = timings.get("sent_matches", "?")
                skipped_known = (timings.get("skipped_known", 0) or 0) + (timings.get("sent_skipped_known", 0) or 0)
                skip_str = ""
                if timings.get("skipped_class"):
                    parts = [f"{cls!r}×{n}" for cls, n in timings["skipped_class"].items()]
                    skip_str = " skip[" + " ".join(parts) + "]"
                extras = []
                if skipped_known:
                    extras.append(f"known={skipped_known}")
                if timings.get("conv_none"):
                    extras.append("conv=None")
                if timings.get("conv_err"):
                    extras.append(f"conv_err={timings['conv_err']!r}")
                if timings.get("sent_err"):
                    extras.append(f"sent_err={timings['sent_err']!r}")
                if timings.get("seed_fallback"):
                    extras.append("seed_fallback")
                extra_str = " [" + " ".join(extras) + "]" if extras else ""
                print(
                    f"  [{idx:3d}/{len(convs_to_process)}] {label} | "
                    f"{elapsed:.3f}s | "
                    f"table={table_rows} sent={sent_matches} new={len(email_objs)}"
                    f"{skip_str}{extra_str} | "
                    f"{subject[:50]}"
                )

            except Exception as e:
                print(f"  [{idx:3d}/{len(convs_to_process)}] ERROR: {e} | {subject[:50]}")
                continue

        # Deduplicate across conversations by outlook_id
        seen = {}
        for obj in to_bulk_create:
            if obj.outlook_id not in seen:
                seen[obj.outlook_id] = obj
        unique_emails = list(seen.values())
        dupes = len(to_bulk_create) - len(unique_emails)
        if dupes:
            print(f"  Deduped {dupes} cross-conversation duplicates ({len(unique_emails)} unique)")

        t0 = time.perf_counter()
        inserted = 0
        skipped = 0
        errors = 0
        for obj in unique_emails:
            try:
                obj.save()
                inserted += 1
            except Exception as e:
                err_str = str(e)
                if "UNIQUE constraint" in err_str:
                    skipped += 1
                else:
                    print(f"    SAVE ERROR outlook_id={obj.outlook_id!r} received_at={obj.received_at!r}: {e}")
                    errors += 1
        _t(f"save ({inserted} new, {skipped} already existed, {errors} errors of {len(unique_emails)} unique)", t0)

        total = time.perf_counter() - sync_start
        print(f"=== Sync done in {total:.2f}s — {new_tickets} new tickets, {inserted} new emails ===\n")
        return new_tickets, inserted

    finally:
        pythoncom.CoUninitialize()
