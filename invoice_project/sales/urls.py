from django.urls import path
from . import views


urlpatterns = [
    path("upload/", views.upload_sales_data, name="upload_sales_data"),
]
