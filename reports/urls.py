from django.urls import path
from . import views

urlpatterns = [
    path("", views.upload_view, name="report_upload"),
    path("pdf/", views.report_pdf, name="report_pdf"),
]