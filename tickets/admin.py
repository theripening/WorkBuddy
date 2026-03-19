from django.contrib import admin
from .models import Assignee, Ticket, TicketEmail, TodoItem, WaitingOn


@admin.register(Assignee)
class AssigneeAdmin(admin.ModelAdmin):
    list_display = ["name", "email"]
    search_fields = ["name", "email"]


@admin.register(Ticket)
class TicketAdmin(admin.ModelAdmin):
    list_display = ["subject", "priority", "assignee", "status", "updated_at"]
    list_filter = ["status", "assignee", "priority"]
    search_fields = ["subject", "notes"]
    list_editable = ["status", "priority"]


@admin.register(TicketEmail)
class TicketEmailAdmin(admin.ModelAdmin):
    list_display = ["subject", "sender", "received_at", "ticket", "conversation_id"]
    list_filter = ["ticket"]
    search_fields = ["subject", "sender", "conversation_id"]
    readonly_fields = ["outlook_id", "conversation_id"]


@admin.register(TodoItem)
class TodoItemAdmin(admin.ModelAdmin):
    list_display = ["what", "ticket", "assignee", "due_date", "done", "done_at"]
    list_filter = ["done", "assignee"]
    search_fields = ["what"]


@admin.register(WaitingOn)
class WaitingOnAdmin(admin.ModelAdmin):
    list_display = ["what", "from_who", "ticket", "asked_date", "expected_date", "resolved"]
    list_filter = ["resolved"]
    search_fields = ["what", "from_who"]
