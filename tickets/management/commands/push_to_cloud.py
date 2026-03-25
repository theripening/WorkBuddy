"""
Management command: push all assigned tickets (+ their open todos and waiting-on
items) to WorkBuddyCloud.

Usage:
    python manage.py push_to_cloud
    python manage.py push_to_cloud --dry-run
"""
from django.core.management.base import BaseCommand

from tickets.cloud import push_ticket, push_todo, push_waiting
from tickets.models import Ticket


class Command(BaseCommand):
    help = "Push all assigned tickets and their open todos/waiting-on items to WorkBuddyCloud."

    def add_arguments(self, parser):
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="Print what would be pushed without actually calling the cloud.",
        )

    def handle(self, *args, **options):
        dry_run = options["dry_run"]

        tickets = (
            Ticket.objects
            .filter(assignee__email__isnull=False)
            .exclude(assignee__email="")
            .exclude(status="completed")
            .prefetch_related("todos", "waiting_on", "emails")
            .select_related("assignee")
        )

        self.stdout.write(f"Found {tickets.count()} assigned open ticket(s).\n")

        pushed_tickets = pushed_todos = pushed_waiting = 0

        for ticket in tickets:
            self.stdout.write(f"  #{ticket.pk}: {ticket.subject[:60]} → {ticket.assignee.email}")
            if not dry_run:
                push_ticket(ticket, ticket.assignee.email)
            pushed_tickets += 1

            for todo in ticket.todos.filter(done=False):
                self.stdout.write(f"    todo: {todo.what[:60]}")
                if not dry_run:
                    push_todo(ticket, todo)
                pushed_todos += 1

            for w in ticket.waiting_on.filter(resolved=False):
                self.stdout.write(f"    waiting: {w.what[:60]}")
                if not dry_run:
                    push_waiting(ticket, w)
                pushed_waiting += 1

        prefix = "[DRY RUN] Would push" if dry_run else "Pushed"
        self.stdout.write(
            f"\n{prefix}: {pushed_tickets} ticket(s), "
            f"{pushed_todos} todo(s), {pushed_waiting} waiting-on item(s)."
        )
