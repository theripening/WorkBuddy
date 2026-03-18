from django.contrib import admin
from .models import Ticket, TicketEmail


@admin.register(Ticket)
class TicketAdmin(admin.ModelAdmin):
    list_display = ["subject", "assignee", "status", "updated_at"]
    list_filter = ["status", "assignee"]
    search_fields = ["subject", "notes"]
    list_editable = ["assignee", "status"]


@admin.register(TicketEmail)
class TicketEmailAdmin(admin.ModelAdmin):
    list_display = ["subject", "sender", "received_at", "ticket", "conversation_id"]
    list_filter = ["ticket"]
    search_fields = ["subject", "sender", "conversation_id"]
    readonly_fields = ["outlook_id", "conversation_id"]
