from django.core.management.base import BaseCommand
from tickets.sync import sync_flagged_emails


class Command(BaseCommand):
    help = "Sync flagged emails from Outlook into the ticket database"

    def handle(self, *args, **options):
        try:
            count = sync_flagged_emails()
            self.stdout.write(self.style.SUCCESS(f"Synced {count} new ticket(s)."))
        except Exception as e:
            self.stderr.write(self.style.ERROR(f"Sync failed: {e}"))
