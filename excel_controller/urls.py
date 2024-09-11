from django.urls import path
from . import views

urlpatterns = [
    path('open_workbook/<path:file_path>/', views.open_workbook, name='open_workbook'),
    path('bring_to_front/', views.bring_to_front, name="bring_to_front"),
    path('close/', views.close_workbook, name='close_workbook'),
    path('next_worksheet/', views.next_worksheet, name='next_worksheet'),
    path('previous_worksheet/', views.previous_worksheet, name='previous_worksheet'),
    path('zoom_in/', views.zoom_in, name='zoom_in'),
    path('zoom_out/', views.zoom_out, name='zoom_out'),
    path('scroll_up/', views.scroll_up, name='scroll_up'),
    path('scroll_down/', views.scroll_down, name='scroll_down'),
    path('scroll_left/', views.scroll_left, name='scroll_left'),
    path('scroll_right/', views.scroll_right, name='scroll_right'),
    path('controls/', views.excel_controls, name='excel_controls'),
]
