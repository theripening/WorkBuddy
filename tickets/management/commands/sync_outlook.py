from django.core.management.base import BaseCommand
from tickets.sync import sync_tracked_folder


class Command(BaseCommand):
    help = "Sync emails from Outlook Tracked folder into the ticket database"

    def handle(self, *args, **options):
        try:
            count = sync_tracked_folder()
            self.stdout.write(self.style.SUCCESS(f"Synced {count} new ticket(s)."))
        except Exception as e:
            self.stderr.write(self.style.ERROR(f"Sync failed: {e}"))
