from django.http import JsonResponse
from django.views.decorators.http import require_GET
from urllib.parse import unquote
from word_controller.controllers.word_controller import WordController

word_controller = WordController()

@require_GET
def open_document(request, file_path):
    decoded_file_path = unquote(file_path)
    try:
        success, message = word_controller.open_document(decoded_file_path)
        status = 200 if success else 500
    except Exception as e:
        print(f"Error opening document: {e}")
        return JsonResponse({"message": f"Failed to open document: {str(e)}"}, status=500)

    return JsonResponse({"message": message}, status=status)

@require_GET
def bring_to_front(request):
    success, message = word_controller.bring_to_front()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def close_document(request):
    success, message = word_controller.close_document()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def scroll_up(request):
    success, message = word_controller.scroll_up()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def scroll_down(request):
    success, message = word_controller.scroll_down()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def zoom_in(request):
    success, message = word_controller.zoom_in()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def zoom_out(request):
    success, message = word_controller.zoom_out()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def enable_read_mode(request):
    success, message = word_controller.enable_read_mode()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def disable_read_mode(request):
    success, message = word_controller.disable_read_mode()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def word_controls(request):
    return JsonResponse({"message": "Word controls loaded", "status": "success"})
