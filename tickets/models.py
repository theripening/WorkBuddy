from django.db import models

ASSIGNEES = [
    ("", "Me"),
    ("ma", "MA"),
    ("tb", "TB"),
]

STATUS_CHOICES = [
    ("open", "Open"),
    ("in_progress", "In Progress"),
    ("closed", "Closed"),
]


class Ticket(models.Model):
    subject = models.CharField(max_length=500)
    assignee = models.CharField(max_length=50, choices=ASSIGNEES, blank=True, default="")
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="open")
    notes = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-updated_at"]

    def __str__(self):
        return self.subject

    def latest_email(self):
        return self.emails.order_by("-received_at").first()


class TicketEmail(models.Model):
    ticket = models.ForeignKey(Ticket, related_name="emails", on_delete=models.CASCADE)
    outlook_id = models.CharField(max_length=500, unique=True)
    conversation_id = models.CharField(max_length=500, db_index=True)
    subject = models.CharField(max_length=500)
    sender = models.CharField(max_length=255)
    received_at = models.DateTimeField()
    body_preview = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-received_at"]

    def __str__(self):
        return f"{self.sender}: {self.subject}"
