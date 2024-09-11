from django.urls import path
from . import views
from .views import get_skype_contacts

urlpatterns = [
    path('open/', views.open_skype, name='open_skype'),
    path('close/', views.close_skype, name='close_skype'),
    path('call/<str:username>/', views.call_skype_user, name='call_skype_user'),
    path('message/', views.send_message, name='send_message'),
    path('get_contacts/', get_skype_contacts, name='get_skype_contacts'),
]
