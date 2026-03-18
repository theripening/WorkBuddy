from django.shortcuts import render, redirect, get_object_or_404
from django.views.decorators.http import require_POST
from django.contrib import messages
from django.db.models import Q

from .models import Ticket, TicketEmail, ASSIGNEES, STATUS_CHOICES


def ticket_list(request):
    all_tickets = Ticket.objects.all()

    open_count = all_tickets.filter(status="open").count()
    in_progress_count = all_tickets.filter(status="in_progress").count()
    closed_count = all_tickets.filter(status="closed").count()

    tickets = all_tickets
    status_filter = request.GET.get("status", "")
    assignee_filter = request.GET.get("assignee", "")
    search = request.GET.get("q", "")

    if status_filter:
        tickets = tickets.filter(status=status_filter)
    if assignee_filter:
        tickets = tickets.filter(assignee=assignee_filter)
    if search:
        tickets = tickets.filter(
            Q(subject__icontains=search)
            | Q(emails__sender__icontains=search)
        ).distinct()

    # Annotate each ticket with its latest email for preview
    ticket_list_data = []
    for ticket in tickets:
        latest = ticket.latest_email()
        ticket_list_data.append({"ticket": ticket, "latest": latest})

    return render(request, "tickets/ticket_list.html", {
        "ticket_list_data": ticket_list_data,
        "assignees": ASSIGNEES,
        "status_choices": STATUS_CHOICES,
        "current_status": status_filter,
        "current_assignee": assignee_filter,
        "search": search,
        "open_count": open_count,
        "in_progress_count": in_progress_count,
        "closed_count": closed_count,
    })


def ticket_detail(request, pk):
    ticket = get_object_or_404(Ticket, pk=pk)

    # Group emails by conversation_id, sorted newest-first within each group
    emails = ticket.emails.order_by("conversation_id", "-received_at")

    threads = {}
    for email in emails:
        cid = email.conversation_id
        if cid not in threads:
            threads[cid] = []
        threads[cid].append(email)

    # Sort threads by the received_at of their latest email (newest thread first)
    sorted_threads = sorted(threads.values(), key=lambda t: t[0].received_at, reverse=True)

    other_tickets = Ticket.objects.exclude(pk=pk).order_by("subject")

    return render(request, "tickets/ticket_detail.html", {
        "ticket": ticket,
        "threads": sorted_threads,
        "other_tickets": other_tickets,
        "assignees": ASSIGNEES,
        "status_choices": STATUS_CHOICES,
    })


@require_POST
def ticket_update(request, pk):
    ticket = get_object_or_404(Ticket, pk=pk)
    ticket.subject = request.POST.get("subject", ticket.subject) or ticket.subject
    ticket.assignee = request.POST.get("assignee", ticket.assignee)
    ticket.status = request.POST.get("status", ticket.status)
    ticket.notes = request.POST.get("notes", ticket.notes)
    ticket.save()
    return redirect(request.POST.get("next", "tickets:list"))


@require_POST
def ticket_delete(request, pk):
    ticket = get_object_or_404(Ticket, pk=pk)
    ticket.delete()
    return redirect("tickets:list")


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
    messages.success(request, f"Merged {email_count} email(s) from \"{source.subject}\" into this ticket.")
    return redirect("tickets:detail", pk=pk)


@require_POST
def sync_outlook(request):
    try:
        from .sync import sync_tracked_folder
        new_tickets, new_emails = sync_tracked_folder()
        messages.success(
            request,
            f"Sync complete — {new_tickets} new ticket(s), {new_emails} new email(s)."
        )
    except Exception as e:
        messages.error(request, f"Sync failed: {e}")
    return redirect("tickets:list")
