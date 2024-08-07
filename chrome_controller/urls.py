# chrome_controller/urls.py
from django.urls import path
from . import views

urlpatterns = [
    path('open_chrome/', views.open_chrome, name='open_chrome'),
    path('home/', views.go_home, name='go_home'),
    path('back/', views.go_back, name='go_back'),
    path('forward/', views.go_forward, name='go_forward'),
    path('close/', views.close_chrome, name='close_chrome'),
    # Add similar paths for other actions like zooming, scrolling, and tab management
]
