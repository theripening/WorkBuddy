from django.urls import path
from . import views

app_name = "scripttools"

urlpatterns = [
    path("scripts/", views.hub, name="hub"),
    path("scripts/pr/run/", views.pr_run, name="pr_run"),
    path("scripts/findphrase/run/", views.findphrase_run, name="findphrase_run"),
    path("scripts/mergecsv/run/", views.mergecsv_run, name="mergecsv_run"),
    path("scripts/sqlimport/run/", views.sqlimport_run, name="sqlimport_run"),
    path("scripts/searchbig/run/", views.searchbig_run, name="searchbig_run"),
    path("scripts/sqlimport-file/run/", views.sqlimport_file_run, name="sqlimport_file_run"),
]
