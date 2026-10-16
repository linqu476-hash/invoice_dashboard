from django.urls import path

from . import views

urlpatterns = [
    path("", views.index, name="index"),
    path("job/<str:job_id>/", views.job, name="job"),
    path("job/<str:job_id>/download/<str:name>", views.download, name="download"),
    path("job/<str:job_id>/delete/", views.delete_job, name="delete_job"),
    path("master/", views.master_tool, name="master_tool"),
    path("master/<str:mid>/download/", views.master_download, name="master_download"),
    path("consignees/", views.consignees, name="consignees"),
    path("consignees/new/", views.consignee_edit, name="consignee_new"),
    path("consignees/<int:pk>/", views.consignee_edit, name="consignee_edit"),
    path("consignees/<int:pk>/delete/", views.consignee_delete, name="consignee_delete"),
]
