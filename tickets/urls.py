from django.urls import path
from . import views

app_name = "tickets"

urlpatterns = [
    path("", views.dashboard, name="list"),
    path("new/", views.ticket_create, name="create"),
    path("<int:pk>/", views.ticket_detail, name="detail"),
    path("<int:pk>/update/", views.ticket_update, name="update"),
    path("<int:pk>/delete/", views.ticket_delete, name="delete"),
    path("<int:pk>/merge/", views.ticket_merge, name="merge"),
    path("<int:pk>/todo/add/", views.todo_add, name="todo_add"),
    path("todo/<int:item_pk>/done/", views.todo_done, name="todo_done"),
    path("todo/<int:item_pk>/update/", views.todo_update, name="todo_update"),
    path("<int:pk>/waiting-on/add/", views.waiting_on_add, name="waiting_on_add"),
    path("waiting-on/<int:item_pk>/resolve/", views.waiting_on_resolve, name="waiting_on_resolve"),
    path("waiting-on/<int:item_pk>/update/", views.waiting_on_update, name="waiting_on_update"),
    path("email/<int:email_pk>/open/", views.open_in_outlook, name="open_in_outlook"),
    path("sync/", views.sync_outlook, name="sync"),
    path("<int:pk>/sync/", views.sync_ticket, name="sync_ticket"),
    path("<int:pk>/notify/ticket/", views.notify_ticket, name="notify_ticket"),
    path("todo/<int:item_pk>/notify/", views.notify_todo, name="notify_todo"),
]
