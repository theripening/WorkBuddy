from datetime import timedelta

from django.conf import settings
from django.shortcuts import render, redirect, get_object_or_404
from django.views.decorators.http import require_POST
from django.contrib import messages
from django.db.models import Q
from django.utils import timezone

from .models import Ticket, TicketEmail, TodoItem, WaitingOn, Assignee, STATUS_CHOICES, PRIORITY_CHOICES, PRIORITY_ORDER


def _priority_sort(p):
    return PRIORITY_ORDER.get(p, 3)


def _ticket_row(t, assignee_email=None):
    """Build the enriched row dict used by the Work tab split-panel."""
    subject = t.subject or ""
    if " | " in subject:
        parts = subject.split(" | ", 1)
        client_code = parts[0].strip()
        subject_display = parts[1].strip()
    else:
        client_code = ""
        subject_display = subject
    latest = t.latest_email()
    open_todo_count = sum(1 for x in t.todos.all() if not x.done)
    open_waiting_count = sum(1 for x in t.waiting_on.all() if not x.resolved)
    has_reply = bool(
        latest and assignee_email
        and latest.sender
        and latest.sender.lower() != assignee_email.lower()
    )
    all_emails = list(t.emails.all())
    recent_emails = sorted(all_emails, key=lambda e: e.received_at, reverse=True)[:5]
    return {
        "ticket": t,
        "latest": latest,
        "client_code": client_code,
        "subject_display": subject_display,
        "open_todo_count": open_todo_count,
        "open_waiting_count": open_waiting_count,
        "has_reply": has_reply,
        "recent_emails": recent_emails,
        "email_count": len(all_emails),
    }


def dashboard(request):
    today = timezone.now().date()
    stale_days = getattr(settings, "STALE_DAYS", 7)
    stale_cutoff = today - timedelta(days=stale_days)

    all_tickets = Ticket.objects.prefetch_related("emails", "waiting_on", "todos", "todos__waiting_on")

    # --- Stat card counts ---
    overdue_count = TodoItem.objects.filter(done=False, due_date__lt=today).count()
    waiting_count = WaitingOn.objects.filter(resolved=False).count()
    

    # --- TO-DO tab: undone TodoItems across all tickets ---
    # Sort: items with a due date first (ascending), then no-due-date items.
    # Within same due date (or both no date), sort by ticket priority.
    todo_items_qs = (
        TodoItem.objects
        .filter(done=False)
        .select_related("ticket", "thread_email", "assignee", "waiting_on")
    )
    todo_items = []
    for item in todo_items_qs:
        overdue = item.due_date and item.due_date < today
        blocked = bool(item.waiting_on_id and not item.waiting_on.resolved)
        todo_items.append({
            "item": item,
            "overdue": overdue,
            "blocked": blocked,
            "sort_key": (
                1 if blocked else 0,                                # blocked items last
                0 if item.due_date else 1,                          # dated items first
                item.due_date.toordinal() if item.due_date else 0,  # earlier dates first
                _priority_sort(item.ticket.priority),               # higher priority first
            ),
        })
    todo_items.sort(key=lambda x: x["sort_key"])
    my_todos = [x for x in todo_items if not x["item"].assignee_id]
    team_todos = [x for x in todo_items if x["item"].assignee_id]

    # --- STALE tab ---
    stale_tickets = []
    for t in all_tickets.filter(status__in=["created", "acknowledged", "in_progress"]):
        last_activity = t.last_activity_date()
        if last_activity <= stale_cutoff:
            days = (today - last_activity).days
            stale_tickets.append({
                "ticket": t,
                "latest": t.latest_email(),
                "last_activity": last_activity,
                "days_stale": days,
            })
    stale_tickets.sort(key=lambda x: x["days_stale"], reverse=True)
    stale_count = len(stale_tickets)

    # Shared filter params — used by both Triage and All tabs
    status_filter = request.GET.get("status", "")
    assignee_filter = request.GET.get("assignee", "")
    priority_filter = request.GET.get("priority", "")
    search = request.GET.get("q", "")

    # --- NEEDS TRIAGE tab: created tickets with no assignee ---
    triage_qs = all_tickets.filter(status="created", assignee__isnull=True)
    if search:
        triage_qs = triage_qs.filter(
            Q(subject__icontains=search) | Q(emails__sender__icontains=search)
        ).distinct()
    if priority_filter:
        triage_qs = triage_qs.filter(priority=priority_filter)
    triage_tickets = [{"ticket": t, "latest": t.latest_email()} for t in triage_qs]
    triage_tickets.sort(key=lambda x: _priority_sort(x["ticket"].priority))

    # --- RECENT tab ---
    recent_cutoff = today - timedelta(days=stale_days)
    recent_tickets = []
    seen_recent = set()
    for email in TicketEmail.objects.filter(received_at__date__gte=recent_cutoff).order_by("-received_at").select_related("ticket__assignee"):
        if email.ticket_id not in seen_recent:
            seen_recent.add(email.ticket_id)
            recent_tickets.append({"ticket": email.ticket, "latest": email})
    _recent_mine, _recent_assigned = [], {}
    for row in recent_tickets:
        a = row["ticket"].assignee
        if a:
            _recent_assigned.setdefault(a.name, []).append(row)
        else:
            _recent_mine.append(row)
    recent_groups = [{"label": "Mine", "tickets": _recent_mine}] + [
        {"label": name, "tickets": rows}
        for name, rows in sorted(_recent_assigned.items())
    ]

    # --- WAITING ON tab ---
    waiting_items = WaitingOn.objects.filter(resolved=False).select_related("ticket", "thread_email").order_by("expected_date", "asked_date")
    waiting_on_list = []
    for w in waiting_items:
        waiting_on_list.append({
            "item": w,
            "overdue": w.expected_date and w.expected_date < today,
            "days_waiting": w.days_waiting(),
            "blocking_count": w.dependent_todos.filter(done=False).count(),
        })

    # --- OPEN tab ---
    _open_qs = all_tickets.filter(status__in=["created", "acknowledged", "in_progress"])
    my_open_tickets = [
        _ticket_row(t)
        for t in _open_qs.filter(assignee__isnull=True).order_by("subject")
    ]
    assigned_open_tickets = sorted(
        [_ticket_row(t, t.assignee.email) for t in _open_qs.filter(assignee__isnull=False)],
        key=lambda x: (x["ticket"].assignee.name.lower(), x["ticket"].subject.lower() if x["ticket"].subject else ""),
    )

    open_count = len(my_open_tickets)
    team_open_count = len(assigned_open_tickets)

    # --- WORK tab: ticket list grouped by person (todos live inside the panel) ---
    from collections import defaultdict
    _team_tickets_by = defaultdict(list)
    for row in assigned_open_tickets:
        _team_tickets_by[row["ticket"].assignee.name].append(row)
    work_groups_list = [{"label": "Mine", "tickets": my_open_tickets}] + [
        {"label": name, "tickets": _team_tickets_by[name]}
        for name in sorted(_team_tickets_by.keys())
    ]
    work_ticket_total = len(my_open_tickets) + len(assigned_open_tickets)


    # --- ALL tab ---
    all_tab = all_tickets
    if status_filter:
        all_tab = all_tab.filter(status=status_filter)
    if assignee_filter:
        all_tab = all_tab.filter(assignee_id=assignee_filter)
    if priority_filter:
        all_tab = all_tab.filter(priority=priority_filter)
    if search:
        all_tab = all_tab.filter(
            Q(subject__icontains=search) | Q(emails__sender__icontains=search)
        ).distinct()
    all_tab_data = [{"ticket": t, "latest": t.latest_email()} for t in all_tab]

    work_pks = {x["ticket"].pk for x in my_open_tickets} | {x["ticket"].pk for x in assigned_open_tickets}
    form_pks = {x["ticket"].pk for x in all_tab_data} | {x["ticket"].pk for x in triage_tickets} | work_pks
    form_tickets = Ticket.objects.filter(pk__in=form_pks).select_related("assignee")

    active_tab = request.GET.get("tab", "todo")
    assignees = Assignee.objects.all()

    return render(request, "tickets/dashboard.html", {
        "today": today,
        "stale_days": stale_days,
        "active_tab": active_tab,
        "overdue_count": overdue_count,
        "stale_count": stale_count,
        "waiting_count": waiting_count,
        "open_count": open_count,
        "team_open_count": team_open_count,
        "work_groups_list": work_groups_list,
        "work_ticket_total": work_ticket_total,
        "my_todos": my_todos,
        "team_todos": team_todos,
        "stale_tickets": stale_tickets,
        "triage_tickets": triage_tickets,
        "my_open_tickets": my_open_tickets,
        "assigned_open_tickets": assigned_open_tickets,
        "recent_tickets": recent_tickets,
        "recent_groups": recent_groups,
        "waiting_on_list": waiting_on_list,
        "all_tab_data": all_tab_data,
        "form_tickets": form_tickets,
        "assignees": assignees,
        "status_choices": STATUS_CHOICES,
        "priority_choices": PRIORITY_CHOICES,
        "current_status": status_filter,
        "current_assignee": assignee_filter,
        "current_priority": priority_filter,
        "search": search,
    })


def ticket_detail(request, pk):
    ticket = get_object_or_404(Ticket.objects.select_related("assignee"), pk=pk)

    emails = ticket.emails.order_by("conversation_id", "-received_at")
    threads = {}
    for email in emails:
        threads.setdefault(email.conversation_id, []).append(email)
    sorted_threads = sorted(threads.values(), key=lambda t: t[0].received_at, reverse=True)

    other_tickets = Ticket.objects.exclude(pk=pk).order_by("subject")
    waiting_on = ticket.waiting_on.all()
    unresolved_waiting_on = ticket.waiting_on.filter(resolved=False)
    todos = ticket.todos.select_related("assignee", "thread_email", "waiting_on").all()
    assignees = Assignee.objects.all()

    return render(request, "tickets/ticket_detail.html", {
        "ticket": ticket,
        "threads": sorted_threads,
        "other_tickets": other_tickets,
        "waiting_on": waiting_on,
        "unresolved_waiting_on": unresolved_waiting_on,
        "todos": todos,
        "assignees": assignees,
        "status_choices": STATUS_CHOICES,
        "priority_choices": PRIORITY_CHOICES,
        "today": timezone.now().date(),
    })


def ticket_create(request):
    if request.method == "POST":
        subject = request.POST.get("subject", "").strip() or "Untitled"
        ticket = Ticket.objects.create(subject=subject)
        messages.success(request, f'Ticket "{ticket.subject}" created.')
        return redirect("tickets:detail", pk=ticket.pk)
    return render(request, "tickets/ticket_create.html")


@require_POST
def ticket_update(request, pk):
    from .sync import unflag_ticket_emails
    from .cloud import forward_to_assignee, push_ticket, push_status
    ticket = get_object_or_404(Ticket, pk=pk)
    prev_subject = ticket.subject
    ticket.subject = request.POST.get("subject", ticket.subject) or ticket.subject
    prev_assignee_id = ticket.assignee_id
    assignee_raw = request.POST.get("assignee", "")
    ticket.assignee_id = int(assignee_raw) if assignee_raw else None
    prev_status = ticket.status
    ticket.status = request.POST.get("status", ticket.status)
    ticket.priority = request.POST.get("priority", ticket.priority)
    ticket.notes = request.POST.get("notes", ticket.notes)
    ticket.save()

    # Assignee changed and is now set — forward email + push to cloud
    if ticket.assignee_id and ticket.assignee_id != prev_assignee_id:
        assignee_email = ticket.assignee.email
        if assignee_email:
            try:
                my_email = forward_to_assignee(ticket, assignee_email)
                if my_email:
                    push_ticket(ticket, assignee_email, my_email)
            except Exception as e:
                messages.warning(request, f"Ticket saved but could not forward to assignee: {e}")

    # Subject changed — push updated subject to cloud
    if ticket.subject != prev_subject and ticket.assignee_id and ticket.assignee and ticket.assignee.email:
        try:
            push_ticket(ticket, ticket.assignee.email)
        except Exception:
            pass

    # Status changed — push to cloud
    if ticket.status != prev_status:
        try:
            push_status(ticket, ticket.status, my_email=None)
        except Exception:
            pass

    # Ticket completed — unflag in Outlook
    if ticket.status == "completed" and prev_status != "completed":
        try:
            cleared = unflag_ticket_emails(ticket)
            if cleared:
                messages.success(request, f"Ticket completed — {cleared} Outlook flag(s) cleared.")
            else:
                messages.warning(request, "Ticket completed but no flagged Outlook emails were found to clear.")
        except Exception as e:
            messages.warning(request, f"Ticket saved but could not unflag Outlook emails: {e}")

    return redirect(request.POST.get("next", "tickets:list"))


@require_POST
def ticket_delete(request, pk):
    ticket = get_object_or_404(Ticket, pk=pk)
    ticket.delete()
    return redirect("tickets:list")


@require_POST
def ticket_merge(request, pk):
    ticket = get_object_or_404(Ticket, pk=pk)
    source_pk = request.POST.get("source_ticket")
    if not source_pk:
        messages.error(request, "No ticket selected to merge.")
        return redirect("tickets:detail", pk=pk)
    source = get_object_or_404(Ticket, pk=source_pk)
    if source.pk == ticket.pk:
        messages.error(request, "Cannot merge a ticket with itself.")
        return redirect("tickets:detail", pk=pk)
    email_count = source.emails.count()
    source.emails.update(ticket=ticket)
    source.delete()
    messages.success(request, f'Merged {email_count} email(s) from "{source.subject}" into this ticket.')
    return redirect("tickets:detail", pk=pk)


@require_POST
def todo_add(request, pk):
    ticket = get_object_or_404(Ticket, pk=pk)
    what = request.POST.get("what", "").strip()
    if what:
        due_raw = request.POST.get("due_date", "").strip()
        assignee_raw = request.POST.get("assignee", "").strip()
        email_pk = request.POST.get("thread_email_pk", "").strip()
        thread_email = TicketEmail.objects.filter(pk=email_pk).first() if email_pk else None
        waiting_on_raw = request.POST.get("waiting_on", "").strip()
        todo = TodoItem.objects.create(
            ticket=ticket,
            thread_email=thread_email,
            assignee_id=int(assignee_raw) if assignee_raw else None,
            what=what,
            due_date=due_raw if due_raw else None,
            waiting_on_id=int(waiting_on_raw) if waiting_on_raw else None,
        )
        if ticket.assignee and ticket.assignee.email:
            from .cloud import push_todo
            push_todo(ticket, todo)
    return redirect("tickets:detail", pk=pk)


@require_POST
def todo_done(request, item_pk):
    item = get_object_or_404(TodoItem.objects.select_related("ticket__assignee"), pk=item_pk)
    item.done = True
    item.done_at = timezone.now()
    item.save()
    if item.ticket.assignee and item.ticket.assignee.email:
        from .cloud import complete_todo
        complete_todo(item)
    from django.http import HttpResponseRedirect
    next_url = request.POST.get("next") or f"/tickets/{item.ticket_id}/"
    return HttpResponseRedirect(next_url)


@require_POST
def todo_update(request, item_pk):
    item = get_object_or_404(TodoItem, pk=item_pk)
    if "what" in request.POST:
        what = request.POST.get("what", "").strip()
        if what:
            item.what = what
    if "due_date" in request.POST:
        due_raw = request.POST.get("due_date", "").strip()
        item.due_date = due_raw if due_raw else None
    if "assignee" in request.POST:
        assignee_raw = request.POST.get("assignee", "").strip()
        item.assignee_id = int(assignee_raw) if assignee_raw else None
    if "waiting_on" in request.POST:
        waiting_on_raw = request.POST.get("waiting_on", "").strip()
        item.waiting_on_id = int(waiting_on_raw) if waiting_on_raw else None
    item.save()
    from django.http import HttpResponseRedirect
    return HttpResponseRedirect(request.POST.get("next", "/"))


@require_POST
def waiting_on_add(request, pk):
    ticket = get_object_or_404(Ticket, pk=pk)
    what = request.POST.get("what", "").strip()
    if what:
        from_who = request.POST.get("from_who", "").strip()
        expected_raw = request.POST.get("expected_date", "").strip()
        email_pk = request.POST.get("thread_email_pk", "").strip()
        thread_email = TicketEmail.objects.filter(pk=email_pk).first() if email_pk else None
        waiting = WaitingOn.objects.create(
            ticket=ticket,
            thread_email=thread_email,
            what=what,
            from_who=from_who,
            expected_date=expected_raw if expected_raw else None,
        )
        if ticket.assignee and ticket.assignee.email:
            from .cloud import push_waiting
            push_waiting(ticket, waiting)
    return redirect("tickets:detail", pk=pk)


@require_POST
def waiting_on_resolve(request, item_pk):
    item = get_object_or_404(WaitingOn.objects.select_related("ticket__assignee"), pk=item_pk)
    item.resolved = True
    item.resolved_at = timezone.now().date()
    item.save()
    if item.ticket.assignee and item.ticket.assignee.email:
        from .cloud import resolve_waiting
        resolve_waiting(item)
    return redirect("tickets:detail", pk=item.ticket_id)


@require_POST
def waiting_on_update(request, item_pk):
    item = get_object_or_404(WaitingOn, pk=item_pk)
    what = request.POST.get("what", "").strip()
    if what:
        item.what = what
    item.from_who = request.POST.get("from_who", "").strip()
    expected_raw = request.POST.get("expected_date", "").strip()
    item.expected_date = expected_raw if expected_raw else None
    item.save()
    return redirect("tickets:detail", pk=item.ticket_id)


@require_POST
def open_in_outlook(request, email_pk):
    email = get_object_or_404(TicketEmail, pk=email_pk)
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            item = namespace.GetItemFromID(email.outlook_id)
            item.Display()
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        messages.error(request, f"Could not open in Outlook: {e}")
    return redirect(request.POST.get("next", "tickets:list"))


@require_POST
def sync_outlook(request):
    try:
        from .sync import sync_flagged_emails
        new_tickets, new_emails = sync_flagged_emails()
        messages.success(request, f"Sync complete — {new_tickets} new ticket(s), {new_emails} new email(s).")
    except Exception as e:
        messages.error(request, f"Sync failed: {e}")
        return redirect("tickets:list")

    # Pull cloud subject edits first, then push local state back out
    try:
        from .cloud import pull_subjects_from_cloud, push_ticket
        pull_subjects_from_cloud()
        assigned = (
            Ticket.objects
            .filter(assignee__isnull=False)
            .exclude(assignee__email="")
            .exclude(status="completed")
            .select_related("assignee")
        )
        for t in assigned:
            if t.assignee and t.assignee.email:
                push_ticket(t, t.assignee.email)
    except Exception:
        pass

    return redirect("tickets:list")


@require_POST
def sync_new_outlook(request):
    # 1. Pull cloud subject edits into local DB first
    try:
        from .cloud import pull_subjects_from_cloud, push_ticket
        pull_subjects_from_cloud()
    except Exception:
        pass

    # 2. Sync new Outlook tickets
    try:
        from .sync import sync_new_flagged
        new_tickets, new_emails = sync_new_flagged()
        messages.success(request, f"Quick sync complete — {new_tickets} new ticket(s).")
    except Exception as e:
        messages.error(request, f"Quick sync failed: {e}")
        return redirect("tickets:list")

    # 3. Push current local state (including any just-pulled edits) to cloud
    try:
        assigned = (
            Ticket.objects
            .filter(assignee__isnull=False)
            .exclude(assignee__email="")
            .exclude(status="completed")
            .select_related("assignee")
        )
        for t in assigned:
            if t.assignee and t.assignee.email:
                push_ticket(t, t.assignee.email)
    except Exception:
        pass

    return redirect("tickets:list")


@require_POST
def sync_ticket(request, pk):
    ticket = get_object_or_404(Ticket, pk=pk)
    try:
        from .sync import sync_ticket_conversations
        new_emails = sync_ticket_conversations(ticket)
        messages.success(request, f"Sync complete — {new_emails} new email(s).")
    except Exception as e:
        messages.error(request, f"Sync failed: {e}")
    return redirect("tickets:detail", pk=pk)


@require_POST
def notify_ticket(request, pk):
    """Open Outlook compose addressed to the ticket's assignee with ticket details and latest email attached."""
    ticket = get_object_or_404(Ticket.objects.select_related("assignee"), pk=pk)
    if not ticket.assignee or not ticket.assignee.email:
        messages.error(request, "Ticket assignee has no email address. Add one in the admin.")
        return redirect("tickets:detail", pk=pk)

    body = "\n".join(filter(None, [
        f"I've assigned the following ticket to you.",
        "",
        f"Subject: {ticket.subject}",
        f"Status: {ticket.get_status_display()}",
        f"Priority: {ticket.priority or 'Not set'}",
        ("" if not ticket.notes else f"\nNotes:\n{ticket.notes}"),
    ]))

    seed_email = ticket.latest_email()
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            mail = outlook.CreateItem(0)
            mail.To = ticket.assignee.email
            mail.Subject = f"[WorkBuddy] {ticket.subject}"
            mail.Body = body
            if seed_email:
                original = namespace.GetItemFromID(seed_email.outlook_id)
                mail.Attachments.Add(original)
            mail.Display()
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        messages.error(request, f"Could not open Outlook: {e}")
    return redirect("tickets:detail", pk=pk)


@require_POST
def notify_todo(request, item_pk):
    """Open Outlook compose addressed to the todo's assignee (falls back to ticket assignee)."""
    todo = get_object_or_404(TodoItem.objects.select_related("assignee", "ticket__assignee", "thread_email"), pk=item_pk)
    ticket = todo.ticket
    assignee = todo.assignee or ticket.assignee
    if not assignee or not assignee.email:
        messages.error(request, "No assignee email set. Add one in the admin.")
        return redirect("tickets:detail", pk=ticket.pk)

    body_lines = [
        "I've assigned a to-do item to you.",
        "",
        f"To-do: {todo.what}",
        f"Ticket: {ticket.subject}",
    ]
    if todo.due_date:
        body_lines.append(f"Due: {todo.due_date.strftime('%B %d, %Y')}")

    seed_email = todo.thread_email or ticket.latest_email()
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            mail = outlook.CreateItem(0)
            mail.To = assignee.email
            mail.Subject = f"[WorkBuddy] To-do: {todo.what}"
            mail.Body = "\n".join(body_lines)
            if seed_email:
                original = namespace.GetItemFromID(seed_email.outlook_id)
                mail.Attachments.Add(original)
            mail.Display()
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        messages.error(request, f"Could not open Outlook: {e}")
    return redirect("tickets:detail", pk=ticket.pk)
