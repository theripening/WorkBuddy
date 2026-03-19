from django.urls import path
from . import views

app_name = "tickets"

urlpatterns = [
    path("", views.ticket_list, name="list"),
    path("new/", views.ticket_create, name="create"),
    path("<int:pk>/", views.ticket_detail, name="detail"),
    path("<int:pk>/update/", views.ticket_update, name="update"),
    path("<int:pk>/delete/", views.ticket_delete, name="delete"),
    path("<int:pk>/merge/", views.ticket_merge, name="merge"),
    path("email/<int:email_pk>/open/", views.open_in_outlook, name="open_in_outlook"),
    path("sync/", views.sync_outlook, name="sync"),
]
