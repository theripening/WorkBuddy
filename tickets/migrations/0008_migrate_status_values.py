from django.db import migrations


def forwards(apps, schema_editor):
    Ticket = apps.get_model("tickets", "Ticket")
    Ticket.objects.filter(status="open").update(status="created")
    Ticket.objects.filter(status="closed").update(status="completed")


def backwards(apps, schema_editor):
    Ticket = apps.get_model("tickets", "Ticket")
    Ticket.objects.filter(status="created").update(status="open")
    Ticket.objects.filter(status="completed").update(status="closed")


class Migration(migrations.Migration):
    dependencies = [
        ("tickets", "0007_status_choices"),
    ]

    operations = [
        migrations.RunPython(forwards, backwards),
    ]
