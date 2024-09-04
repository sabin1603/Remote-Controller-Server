from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path('admin/', admin.site.urls),
    path('api/file_explorer/', include('file_explorer.urls')),
    path('api/teams/', include('teams_controller.urls')),
    path('api/powerpoint/', include('powerpoint_controller.urls')),
    path('api/excel/', include('excel_controller.urls')),
    path('api/word/', include('word_controller.urls')),
    path('api/pdf-reader/', include('pdf_reader.urls')),
    path('api/chrome/', include('chrome_controller.urls')),
    path('api/skype/', include('skype_controller.urls')),

]
