from django.urls import path
from .views import open_powerpoint

urlpatterns = [
    path('open/', open_powerpoint, name='open_powerpoint'),
]
