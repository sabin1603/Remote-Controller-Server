from django.urls import path
from . import views

urlpatterns = [
    path('list_workbooks/', views.list_workbooks, name='list_workbooks'),
    path('open_workbook/<str:file_name>/', views.open_workbook, name='open_workbook'),
    path('next_worksheet/', views.next_worksheet, name='next_worksheet'),
    path('previous_worksheet/', views.previous_worksheet, name='previous_worksheet'),
    path('zoom_in/', views.zoom_in, name='zoom_in'),
    path('zoom_out/', views.zoom_out, name='zoom_out'),
    path('scroll_up/', views.scroll_up, name='scroll_up'),
    path('scroll_down/', views.scroll_down, name='scroll_down'),
    path('scroll_left/', views.scroll_left, name='scroll_left'),
    path('scroll_right/', views.scroll_right, name='scroll_right'),
]
