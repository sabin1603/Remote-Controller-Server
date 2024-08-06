from django.urls import path
from . import views

urlpatterns = [
    path('list_documents/', views.list_documents, name='list_documents'),
    path('open_document/<str:file_name>/', views.open_document, name='open_document'),
    path('close_document/', views.close_document, name='close_document'),
    path('scroll_up/', views.scroll_up, name='scroll_up'),
    path('scroll_down/', views.scroll_down, name='scroll_down'),
    path('zoom_in/', views.zoom_in, name='zoom_in'),
    path('zoom_out/', views.zoom_out, name='zoom_out'),
    path('enable_read_mode/', views.enable_read_mode, name='enable_read_mode'),
    path('disable_read_mode/', views.disable_read_mode, name='disable_read_mode'),
    path('next_page/', views.next_page, name='next_page'),
    path('previous_page/', views.previous_page, name='previous_page'),
]
