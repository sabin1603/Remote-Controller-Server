from django.http import JsonResponse
from django.views.decorators.http import require_GET
from excel_controller.controllers.excel_controller import ExcelController

excel_controller = ExcelController()

@require_GET
def open_workbook(request, file_path):
    success, message = excel_controller.open_workbook(file_path)
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def bring_to_front(request):
    success, message = excel_controller.bring_to_front()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def close_workbook(request):
    success, message = excel_controller.close_workbook()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def next_worksheet(request):
    success, message = excel_controller.change_worksheet(next_sheet=True)
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def previous_worksheet(request):
    success, message = excel_controller.change_worksheet(next_sheet=False)
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def zoom_in(request):
    success, message = excel_controller.zoom(zoom_in=True)
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def zoom_out(request):
    success, message = excel_controller.zoom(zoom_in=False)
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def scroll_up(request):
    success, message = excel_controller.scroll('up')
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def scroll_down(request):
    success, message = excel_controller.scroll('down')
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def scroll_left(request):
    success, message = excel_controller.scroll('left')
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def scroll_right(request):
    success, message = excel_controller.scroll('right')
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def excel_controls(request):
    return JsonResponse({"message": "Excel controls loaded", "status": "success"})
