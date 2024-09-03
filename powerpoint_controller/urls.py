from django.urls import path
from . import views

urlpatterns = [
    path('open_presentation/', views.open_presentation, name='open_presentation'),
    path('next/', views.next_slide, name='next_slide'),
    path('prev/', views.prev_slide, name='prev_slide'),
    path('start/', views.start_presentation, name='start_presentation'),
    path('end/', views.end_presentation, name='end_presentation'),
    path('controls/', views.powerpoint_controls, name='powerpoint_controls'),
    path('current_presentation/', views.get_current_presentation, name='get_current_presentation'),
]
