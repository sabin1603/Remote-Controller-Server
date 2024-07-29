from django.urls import path
from . import views

urlpatterns = [
    path('list_presentations/', views.list_presentations, name='list_presentations'),
    path('open_presentation/<str:file_name>/', views.open_presentation, name='open_presentation'),
    path('next/', views.next_slide, name='next_slide'),
    path('prev/', views.prev_slide, name='prev_slide'),
    path('start/', views.start_presentation, name='start_presentation'),
    path('end/', views.end_presentation, name='end_presentation'),
]
