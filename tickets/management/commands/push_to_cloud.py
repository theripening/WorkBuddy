"""
Management command: push all assigned tickets (+ their open todos and waiting-on
items) to WorkBuddyCloud.

Usage:
    python manage.py push_to_cloud --email you@company.com

Run this once to catch up tickets that were assigned before cloud sync was enabled.
"""
import pythoncom
import win32com.client

from django.core.management.base import BaseCommand

from tickets.cloud import get_my_email, push_ticket, push_todo, push_waiting
from tickets.models import Ticket


class Command(BaseCommand):
    help = "Push all assigned tickets and their open todos/waiting-on items to WorkBuddyCloud."

    def add_arguments(self, parser):
        parser.add_argument(
            "--email",
            type=str,
            default="",
            help="Your Outlook email (assigner). Omit to read from Outlook automatically.",
        )
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="Print what would be pushed without actually calling the cloud.",
        )

    def handle(self, *args, **options):
        dry_run = options["dry_run"]
        my_email = options["email"].strip()

        if not my_email:
            self.stdout.write("Reading your email from Outlook…")
            pythoncom.CoInitialize()
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                my_email = get_my_email(namespace)
            finally:
                pythoncom.CoUninitialize()

        if not my_email:
            self.stderr.write("Could not determine your email. Pass --email your@address.com")
            return

        self.stdout.write(f"Assigner email: {my_email}")

        tickets = (
            Ticket.objects
            .filter(assignee__email__isnull=False)
            .exclude(assignee__email="")
            .exclude(status="completed")
            .prefetch_related("todos", "waiting_on", "emails", "assignee")
        )

        self.stdout.write(f"Found {tickets.count()} assigned open ticket(s).\n")

        pushed_tickets = 0
        pushed_todos = 0
        pushed_waiting = 0

        for ticket in tickets:
            assignee_email = ticket.assignee.email
            self.stdout.write(f"  Ticket #{ticket.pk}: {ticket.subject[:60]} → {assignee_email}")

            if not dry_run:
                push_ticket(ticket, assignee_email, my_email)
            pushed_tickets += 1

            open_todos = ticket.todos.filter(done=False)
            for todo in open_todos:
                self.stdout.write(f"    todo: {todo.what[:60]}")
                if not dry_run:
                    push_todo(ticket, todo)
                pushed_todos += 1

            open_waiting = ticket.waiting_on.filter(resolved=False)
            for w in open_waiting:
                self.stdout.write(f"    waiting: {w.what[:60]}")
                if not dry_run:
                    push_waiting(ticket, w)
                pushed_waiting += 1

        prefix = "[DRY RUN] Would push" if dry_run else "Pushed"
        self.stdout.write(
            f"\n{prefix}: {pushed_tickets} ticket(s), "
            f"{pushed_todos} todo(s), {pushed_waiting} waiting-on item(s)."
        )
