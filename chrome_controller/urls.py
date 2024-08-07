from django.urls import path
from . import views

urlpatterns = [
    # Chrome control endpoints
    path('open_chrome/', views.open_chrome, name='open_chrome'),
    path('new_tab/', views.new_tab, name='new_tab'),
    path('go_home/', views.go_home, name='go_home'),
    path('go_back/', views.go_back, name='go_back'),
    path('go_forward/', views.go_forward, name='go_forward'),
    path('close_current_tab/', views.close_current_tab, name='close_current_tab'),
    path('zoom_in/', views.zoom_in, name='zoom_in'),
    path('zoom_out/', views.zoom_out, name='zoom_out'),
    path('scroll_up/', views.scroll_up, name='scroll_up'),
    path('scroll_down/', views.scroll_down, name='scroll_down'),
    path('go_to_left_tab/', views.go_to_left_tab, name='go_to_left_tab'),
    path('go_to_right_tab/', views.go_to_right_tab, name='go_to_right_tab'),
    path('close_chrome/', views.close_chrome, name='close_chrome'),
    path('navigate/', views.navigate, name='navigate'),
    path('navigate_up/', views.navigate_up, name='navigate_up'),
    path('navigate_down/', views.navigate_down, name='navigate_down'),
]
