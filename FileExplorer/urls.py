from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path('admin/', admin.site.urls),
    path('api/file_explorer/', include('file_explorer.urls')),
    path('api/teams/', include('teams_controller.urls')),
    path('api/powerpoint/', include('powerpoint_controller.urls')),  # Add this line
]
