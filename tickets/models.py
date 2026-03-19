from django.db import models
from django.utils import timezone


STATUS_CHOICES = [
    ("open", "Open"),
    ("in_progress", "In Progress"),
    ("closed", "Closed"),
]

PRIORITY_CHOICES = [
    ("high", "High"),
    ("medium", "Medium"),
    ("low", "Low"),
]

PRIORITY_ORDER = {"high": 0, "medium": 1, "low": 2, "": 3}


class Assignee(models.Model):
    name = models.CharField(max_length=100)
    email = models.EmailField(max_length=255, blank=True)

    class Meta:
        ordering = ["name"]

    def __str__(self):
        return self.name


class Ticket(models.Model):
    subject = models.CharField(max_length=500)
    assignee = models.ForeignKey(
        Assignee, null=True, blank=True, on_delete=models.SET_NULL, related_name="tickets"
    )
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="open")
    priority = models.CharField(max_length=10, choices=PRIORITY_CHOICES, blank=True, default="")
    notes = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-updated_at"]

    def __str__(self):
        return self.subject

    def latest_email(self):
        return self.emails.order_by("-received_at").first()

    def last_activity_date(self):
        return self.updated_at.date()

    def open_todo_count(self):
        return self.todos.filter(done=False).count()


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


class TodoItem(models.Model):
    ticket = models.ForeignKey(Ticket, related_name="todos", on_delete=models.CASCADE)
    assignee = models.ForeignKey(
        Assignee, null=True, blank=True, on_delete=models.SET_NULL, related_name="todos"
    )
    thread_email = models.ForeignKey(
        TicketEmail, null=True, blank=True, on_delete=models.SET_NULL,
        related_name="todos", help_text="Latest email of the linked thread"
    )
    what = models.CharField(max_length=500)
    due_date = models.DateField(null=True, blank=True)
    done = models.BooleanField(default=False)
    done_at = models.DateTimeField(null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["done", "due_date", "created_at"]

    def __str__(self):
        return self.what


class WaitingOn(models.Model):
    ticket = models.ForeignKey(Ticket, related_name="waiting_on", on_delete=models.CASCADE)
    thread_email = models.ForeignKey(
        TicketEmail, null=True, blank=True, on_delete=models.SET_NULL,
        related_name="waiting_on_items", help_text="Latest email of the linked thread"
    )
    what = models.CharField(max_length=500)
    from_who = models.CharField(max_length=255, blank=True)
    asked_date = models.DateField(default=timezone.now)
    expected_date = models.DateField(null=True, blank=True)
    resolved = models.BooleanField(default=False)
    resolved_at = models.DateField(null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["resolved", "asked_date"]

    def __str__(self):
        return f"{self.what} (from {self.from_who or '?'})"

    def days_waiting(self):
        end = self.resolved_at if self.resolved_at else timezone.now().date()
        return (end - self.asked_date).days
