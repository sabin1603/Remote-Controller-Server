from django.urls import path
from . import views

urlpatterns = [
    # Endpoint to list all PDF files
    path('list/', views.list_pdfs, name='list_pdfs'),

    # Endpoint to open a specific PDF file
    path('open/<str:file_name>/', views.open_pdf, name='open_pdf'),

    # Endpoints for PDF controls
    path('scroll_up/', views.scroll_up, name='scroll_up'),
    path('scroll_down/', views.scroll_down, name='scroll_down'),
    path('zoom_in/', views.zoom_in, name='zoom_in'),
    path('zoom_out/', views.zoom_out, name='zoom_out'),
    path('enable_read_mode/', views.enable_read_mode, name='enable_read_mode'),
    path('disable_read_mode/', views.disable_read_mode, name='disable_read_mode'),
    path('save/', views.save_pdf, name='save_pdf'),
    path('print/', views.print_pdf, name='print_pdf'),
    path('close/', views.close_pdf, name='close_pdf'),
]
