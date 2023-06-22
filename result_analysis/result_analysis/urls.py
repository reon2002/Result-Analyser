"""
URL configuration for result_analysis project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/4.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from . import views
from django.conf import settings
from django.conf.urls.static import static
# from.views import analyze_S4S5

from django.urls import path
urlpatterns = [
    path('admin/', admin.site.urls),
    path('upload/',views.login, name='login'),
    path('upload/upload.html/', views.upload_pdf, name='upload'),
    path('upload/upload.html/download.html/', views.download_pdf, name='download'),
    path('upload/upload.html/download.html/resultexcel/', views.download_excel_view, name='resultexcel'),
    path('upload/upload.html/download.html/resultpdf/', views.download_pdf_view, name='resultpdf'),
    


    # path('download/pdf/', views.download_pdf, name='download_pdf'),
    # path('download/excel/', views.download_excel, name='download_excel')
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)