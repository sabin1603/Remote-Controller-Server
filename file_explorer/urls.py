from django.urls import path
from .views import list_files, open_file, download_file, index

urlpatterns = [
    path('', index, name='index'),
    path('list/', list_files, name='list_files'),
    path('open/', open_file, name='open_file'),
    path('download/', download_file, name='download_file'),
]
