from django.urls import path
from .views import open_teams, get_contacts, start_call

urlpatterns = [
    path('open/', open_teams, name='open_teams'),
    path('contacts/', get_contacts, name='get_contacts'),
    path('api/teams/start_call/', start_call, name='start_call'),
]
