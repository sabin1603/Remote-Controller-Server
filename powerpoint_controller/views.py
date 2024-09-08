from django.http import JsonResponse
from django.views.decorators.http import require_GET
from urllib.parse import unquote
from powerpoint_controller.controllers.powerpoint_controller import PowerPointController

powerpoint_controller = PowerPointController()

@require_GET
def open_presentation(request, file_path):
    decoded_file_path = unquote(file_path)
    try:
        success, message = powerpoint_controller.open_presentation(decoded_file_path)
        status = 200 if success else 500
    except Exception as e:
        print(f"Error opening presentation: {e}")
        return JsonResponse({"message": f"Failed to open presentation: {str(e)}"}, status=500)

    return JsonResponse({"message": message}, status=status)

@require_GET
def close_presentation(request):
    success, message = powerpoint_controller.close_presentation()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def next_slide(request):
    success, message = powerpoint_controller.next_slide()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def prev_slide(request):
    success, message = powerpoint_controller.prev_slide()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def start_presentation(request):
    success, message = powerpoint_controller.start_presentation()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def end_presentation(request):
    success, message = powerpoint_controller.end_presentation()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def powerpoint_controls(request):
    return JsonResponse({"message": "PowerPoint controls loaded", "status": "success"})
