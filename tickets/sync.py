import logging
import time
from collections import Counter
from datetime import datetime

from .models import Ticket, TicketEmail

logger = logging.getLogger(__name__)

# Outlook folder/flag constants
OL_FLAG_MARKED = 2
OL_FLAG_NOT_FLAGGED = 0
OL_FOLDER_TODO = 28

# Columns fetched in a single GetTable call — no body preview here;
# we fetch .Body via GetItemFromID when needed instead of MAPI proptags
# which return unpredictable binary types across Outlook versions.
_TABLE_COLUMNS = ["EntryID", "ConversationID", "MessageClass", "Subject", "SenderName", "ReceivedTime", "SentOn"]


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


def _build_email_from_row(ticket, row, namespace=None, conv_id="", known_outlook_ids=None):
    """
    Build TicketEmail from a conversation table row, fetching body via COM.

    Outlook's GetConversation().GetTable() returns session-temporary short-form
    EntryIDs that change each time Outlook connects. We call GetItemFromID to
    fetch the body anyway, so we use live.EntryID (the stable store-bound ID)
    as the outlook_id. We also do a secondary check against known_outlook_ids
    with the stable ID so we don't re-save emails whose short-form ID changed.

    Returns (TicketEmail | None). None means the email is already known.
    """
    short_id = row["EntryID"]
    stable_id = short_id
    body = ""
    if namespace is not None:
        try:
            live = namespace.GetItemFromID(short_id)
            stable_id = live.EntryID  # normalised, persistent across sessions
            body = (live.Body or "")[:300].strip()
        except Exception as e:
            logger.debug("Body fetch failed for %s: %s", short_id, e)
    # Secondary known-check with stable ID — catches emails whose short-form ID
    # changed since last sync
    if known_outlook_ids is not None and stable_id in known_outlook_ids:
        return None
    received_raw = _row_get(row, "ReceivedTime") or _row_get(row, "SentOn")
    if received_raw is None:
        raise ValueError(f"No ReceivedTime or SentOn for EntryID {short_id}")
    return TicketEmail(
        ticket=ticket,
        outlook_id=stable_id,
        conversation_id=_row_get(row, "ConversationID") or conv_id,
        subject=_row_get(row, "Subject") or "(no subject)",
        sender=_row_get(row, "SenderName") or "",
        received_at=_parse_received_time(received_raw),
        body_preview=body,
    )


def _collect_conversation_emails(namespace, seed_entry_id, ticket, known_outlook_ids=None):
    """
    Return only NEW TicketEmail objects for the conversation containing seed_entry_id.

    Walks the full conversation table every sync to discover new emails, but skips
    GetItemFromID (body fetch) for any EntryID already in known_outlook_ids.
    GetConversation().GetTable() includes sent replies, so no Sent Items scan needed.

    Returns (list of new TicketEmail instances, dict of timing info).
    """
    if known_outlook_ids is None:
        known_outlook_ids = set()
    timings = {}
    emails_by_entry_id = {}  # EntryID -> TicketEmail, deduped in memory

    t0 = time.perf_counter()
    try:
        seed_item = namespace.GetItemFromID(seed_entry_id)
    except Exception as e:
        logger.warning("GetItemFromID failed for seed %s: %s", seed_entry_id, e)
        return [], {"error": "GetItemFromID failed"}
    timings["GetItemFromID"] = time.perf_counter() - t0

    conv_id = seed_item.ConversationID

    try:
        t0 = time.perf_counter()
        conversation = seed_item.GetConversation()
        timings["GetConversation"] = time.perf_counter() - t0

        if conversation is not None:
            t0 = time.perf_counter()
            table = conversation.GetTable()
            table.Columns.RemoveAll()
            for col in _TABLE_COLUMNS:
                table.Columns.Add(col)
            timings["GetTable"] = time.perf_counter() - t0

            # Stream rows directly — avoids building a full list in memory
            t0 = time.perf_counter()
            table_rows = 0
            skipped_class = Counter()
            skipped_known = 0
            while not table.EndOfTable:
                try:
                    row = table.GetNextRow()
                except Exception as e:
                    logger.debug("GetNextRow failed: %s", e)
                    continue
                table_rows += 1
                msg_class = _row_get(row, "MessageClass") or ""
                if msg_class != "IPM.Note":
                    skipped_class[msg_class] += 1
                    continue
                try:
                    entry_id = row["EntryID"]
                    if entry_id in emails_by_entry_id:
                        continue
                    result = _build_email_from_row(ticket, row, namespace, conv_id, known_outlook_ids)
                    if result is None:
                        skipped_known += 1
                        continue
                    emails_by_entry_id[result.outlook_id] = result
                except Exception as e:
                    logger.warning("Skipping table row: %s", e)
            timings["TableWalk"] = time.perf_counter() - t0
            timings["table_rows"] = table_rows
            timings["skipped_class"] = dict(skipped_class)
            timings["skipped_known"] = skipped_known
        else:
            timings["conv_none"] = True

    except Exception as e:
        logger.warning("GetConversation failed for %s: %s", seed_entry_id, e)
        timings["conv_err"] = str(e)

    # Last resort: at minimum store the seed itself
    if not emails_by_entry_id:
        try:
            emails_by_entry_id[seed_entry_id] = _build_email_obj(ticket, seed_item)
            timings["seed_fallback"] = True
        except Exception as e:
            logger.warning("Seed fallback failed: %s", e)

    return list(emails_by_entry_id.values()), timings


def sync_flagged_emails():
    """
    Connect to Outlook and sync all flagged emails (To-Do folder) + their full
    conversations into the DB. Returns (new_tickets, new_emails) counts.

    Only flagged (red flag) emails are used as ticket seeds. Each sync walks
    conversations fresh to catch new replies. Already-known outlook_ids skip
    the COM body fetch. Individual saves catch constraint violations explicitly.
    """
    import pythoncom
    import win32com.client

    sync_start = time.perf_counter()
    logger.info("=== WorkBuddy Sync Start ===")

    pythoncom.CoInitialize()
    try:
        t0 = time.perf_counter()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        logger.info("Outlook connected in %.3fs", time.perf_counter() - t0)

        # Pre-load known conversations, email IDs, and seed outlook_ids from DB
        t0 = time.perf_counter()
        known_convs = {}       # conv_id -> ticket_id
        known_outlook_ids = set()
        conv_to_seed = {}      # conv_id -> seed outlook_id (avoids re-fetch for known convs)
        for outlook_id, conv_id, ticket_id, is_seed in TicketEmail.objects.values_list(
            "outlook_id", "conversation_id", "ticket_id", "is_seed"
        ):
            known_convs[conv_id] = ticket_id
            known_outlook_ids.add(outlook_id)
            if is_seed:
                conv_to_seed[conv_id] = outlook_id
        logger.info(
            "DB pre-load: %d conversations, %d emails, %d seeds in %.3fs",
            len(known_convs), len(known_outlook_ids), len(conv_to_seed), time.perf_counter() - t0,
        )

        # Pre-load all tickets referenced by known conversations in one query
        t0 = time.perf_counter()
        tickets_cache = {t.pk: t for t in Ticket.objects.filter(pk__in=set(known_convs.values()))}
        logger.debug("Tickets pre-loaded: %d in %.3fs", len(tickets_cache), time.perf_counter() - t0)

        # --- Collect flagged emails from the To-Do virtual folder ---
        # olFolderToDo aggregates flagged items from all folders — no recursion needed.
        # GetTable() and FlagStatus in Restrict both fail on this virtual folder, so
        # we Restrict by MessageClass only, then check FlagStatus in Python.
        # GetFirst/GetNext cursor is used instead of Item(i) indexed access —
        # indexed access on COM collections is O(n²); cursor is O(n).
        t0 = time.perf_counter()
        seen_conv_ids = set()
        convs_to_process = []  # list of (conv_id, entry_id, subject)
        seeds_to_heal = []     # outlook_ids that need is_seed=True written back to DB
        todo_scanned = 0
        todo_flagged = 0
        count_seed_hit = 0
        count_seed_heal = 0
        count_new = 0
        try:
            todo_folder = namespace.GetDefaultFolder(OL_FOLDER_TODO)
            mail_todos = todo_folder.Items.Restrict("[MessageClass] = 'IPM.Note'")
            item = mail_todos.GetFirst()
            while item is not None:
                try:
                    todo_scanned += 1
                    if item.FlagStatus != OL_FLAG_MARKED:
                        item = mail_todos.GetNext()
                        continue
                    todo_flagged += 1
                    cid = item.ConversationID
                    if not cid or cid in seen_conv_ids:
                        item = mail_todos.GetNext()
                        continue
                    seen_conv_ids.add(cid)

                    if cid in conv_to_seed:
                        # Known conv, seed on file — no COM call needed
                        entry_id = conv_to_seed[cid]
                        count_seed_hit += 1
                    elif cid in known_convs:
                        # Known conv, seed missing — use item directly and heal DB after loop
                        entry_id = item.EntryID
                        seeds_to_heal.append(entry_id)
                        count_seed_heal += 1
                    else:
                        # Brand new conversation
                        entry_id = item.EntryID
                        count_new += 1

                    convs_to_process.append((cid, entry_id, item.Subject or "(no subject)"))
                except Exception as e:
                    logger.debug("Skipping flagged item: %s", e)
                item = mail_todos.GetNext()
        except Exception as e:
            logger.warning("To-Do folder error: %s", e)

        # Heal missing is_seed flags so next sync skips the re-fetch
        if seeds_to_heal:
            healed = TicketEmail.objects.filter(outlook_id__in=seeds_to_heal).update(is_seed=True)
            logger.info("Healed is_seed on %d email(s)", healed)

        logger.info(
            "Flag scan: %d scanned, %d flagged, %d unique convs "
            "(seed_hit=%d heal=%d new=%d) in %.3fs",
            todo_scanned, todo_flagged, len(convs_to_process),
            count_seed_hit, count_seed_heal, count_new, time.perf_counter() - t0,
        )

        logger.info("Conversations to process: %d", len(convs_to_process))

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
                    # Re-open completed ticket when its seed email is re-flagged
                    if ticket.status == "completed":
                        ticket.status = "created"
                        ticket.save(update_fields=["status"])
                        logger.info("Re-opened completed ticket %d: %s", ticket.pk, (subject or "")[:50])
                else:
                    ticket = Ticket.objects.create(subject=subject)
                    known_convs[conv_id] = ticket.pk
                    tickets_cache[ticket.pk] = ticket
                    new_tickets += 1

                email_objs, timings = _collect_conversation_emails(namespace, entry_id, ticket, known_outlook_ids)

                # Mark the seed email so we can unflag it later when ticket completes
                for obj in email_objs:
                    if obj.outlook_id == entry_id:
                        obj.is_seed = True
                        break
                else:
                    # seed was already known (skipped) — ensure DB record is marked
                    TicketEmail.objects.filter(outlook_id=entry_id).update(is_seed=True)

                to_bulk_create.extend(email_objs)

                elapsed = time.perf_counter() - conv_start
                label = "NEW" if is_new else "sync"
                extras = []
                if timings.get("GetItemFromID"):
                    extras.append(f"seed={timings['GetItemFromID']:.3f}s")
                if timings.get("GetConversation"):
                    extras.append(f"conv={timings['GetConversation']:.3f}s")
                if timings.get("GetTable"):
                    extras.append(f"tbl={timings['GetTable']:.3f}s")
                if timings.get("TableWalk"):
                    extras.append(f"walk={timings['TableWalk']:.3f}s")
                if timings.get("skipped_known"):
                    extras.append(f"known={timings['skipped_known']}")
                if timings.get("skipped_class"):
                    parts = [f"{cls!r}×{n}" for cls, n in timings["skipped_class"].items()]
                    extras.append("skip[" + " ".join(parts) + "]")
                if timings.get("conv_none"):
                    extras.append("conv=None")
                if timings.get("conv_err"):
                    extras.append(f"conv_err={timings['conv_err']!r}")
                if timings.get("seed_fallback"):
                    extras.append("seed_fallback")
                extra_str = " [" + " ".join(extras) + "]" if extras else ""
                logger.info(
                    "[%3d/%d] %s | %.3fs | table=%s new=%d%s | %s",
                    idx, len(convs_to_process), label, elapsed,
                    timings.get("table_rows", "?"), len(email_objs),
                    extra_str, (subject or "")[:50],
                )

            except Exception as e:
                logger.warning("[%3d/%d] ERROR: %s | %s", idx, len(convs_to_process), e, (subject or "")[:50])
                continue

        # Deduplicate across conversations by outlook_id
        seen = {}
        for obj in to_bulk_create:
            if obj.outlook_id not in seen:
                seen[obj.outlook_id] = obj
        unique_emails = list(seen.values())
        dupes = len(to_bulk_create) - len(unique_emails)
        if dupes:
            logger.info("Deduped %d cross-conversation duplicates (%d unique)", dupes, len(unique_emails))

        t0 = time.perf_counter()
        inserted = 0
        skipped = 0
        errors = 0
        for obj in unique_emails:
            try:
                obj.save()
                inserted += 1
            except Exception as e:
                if "UNIQUE constraint" in str(e):
                    skipped += 1
                else:
                    logger.error(
                        "SAVE ERROR outlook_id=%r received_at=%r: %s",
                        obj.outlook_id, obj.received_at, e,
                    )
                    errors += 1
        logger.info(
            "Save: %d new, %d already existed, %d errors of %d unique in %.3fs",
            inserted, skipped, errors, len(unique_emails), time.perf_counter() - t0,
        )

        total = time.perf_counter() - sync_start
        logger.info("=== Sync done in %.2fs — %d new tickets, %d new emails ===", total, new_tickets, inserted)

        # Pull cloud notes if cloud is configured
        try:
            from .cloud import sync_cloud_notes, get_my_email
            my_email = get_my_email(namespace)
            if my_email:
                new_notes = sync_cloud_notes(my_email)
                if new_notes:
                    logger.info("Cloud: %d new note(s) pulled", new_notes)
        except Exception as e:
            logger.warning("Cloud note sync failed: %s", e)

        return new_tickets, inserted

    finally:
        pythoncom.CoUninitialize()


def sync_ticket_conversations(ticket):
    """
    Sync all known conversations for a single ticket.
    Uses stored outlook_ids as seeds — no folder/flag scan needed.
    Returns count of newly inserted emails.
    """
    import pythoncom
    import win32com.client

    # One seed outlook_id per conversation
    conv_seeds = {}
    for outlook_id, conv_id in TicketEmail.objects.filter(ticket=ticket).values_list("outlook_id", "conversation_id"):
        conv_seeds.setdefault(conv_id, outlook_id)

    if not conv_seeds:
        logger.info("Ticket %d has no emails to sync from", ticket.pk)
        return 0

    known_outlook_ids = set(TicketEmail.objects.values_list("outlook_id", flat=True))

    logger.info("Syncing ticket %d: %d conversation(s)", ticket.pk, len(conv_seeds))
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        to_save = []
        for conv_id, seed_outlook_id in conv_seeds.items():
            email_objs, timings = _collect_conversation_emails(namespace, seed_outlook_id, ticket, known_outlook_ids)
            to_save.extend(email_objs)
            logger.info(
                "  conv %s: %d new, %d known",
                conv_id[:16], len(email_objs), timings.get("skipped_known", 0),
            )

        seen = {}
        for obj in to_save:
            if obj.outlook_id not in seen:
                seen[obj.outlook_id] = obj

        inserted = 0
        for obj in seen.values():
            try:
                obj.save()
                inserted += 1
            except Exception as e:
                if "UNIQUE constraint" not in str(e):
                    logger.error("SAVE ERROR outlook_id=%r: %s", obj.outlook_id, e)

        logger.info("Ticket %d sync done: %d new email(s)", ticket.pk, inserted)
        return inserted

    finally:
        pythoncom.CoUninitialize()


def unflag_ticket_emails(ticket):
    """
    Clear the Outlook follow-up flag on all seed emails for the given ticket.
    Called when a ticket is marked as completed.
    Returns the number of emails successfully unflagged.
    """
    seed_ids = list(ticket.emails.filter(is_seed=True).values_list("outlook_id", flat=True))
    if not seed_ids:
        logger.info("Ticket %d: no seed emails to unflag", ticket.pk)
        return 0

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        cleared = 0
        for outlook_id in seed_ids:
            try:
                item = namespace.GetItemFromID(outlook_id)
                if item.FlagStatus != OL_FLAG_NOT_FLAGGED:
                    item.FlagStatus = OL_FLAG_NOT_FLAGGED
                    item.Save()
                    cleared += 1
                    logger.debug("Unflagged outlook_id=%s", outlook_id)
            except Exception as e:
                logger.warning("Could not unflag outlook_id=%s: %s", outlook_id, e)
        logger.info("Ticket %d: unflagged %d/%d seed email(s)", ticket.pk, cleared, len(seed_ids))
        return cleared
    finally:
        pythoncom.CoUninitialize()
