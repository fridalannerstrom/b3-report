from django.urls import path
from . import views
from .views import upload_view, report_pdf_page, report_pdf_download

urlpatterns = [
    path("", upload_view, name="report_upload"),
    path("pdf/page/", report_pdf_page, name="report_pdf_page"),
    path("pdf/download/", report_pdf_download, name="report_pdf_download"),
]