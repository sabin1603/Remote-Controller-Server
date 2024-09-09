from django.http import JsonResponse
from django.views.decorators.http import require_GET
from urllib.parse import unquote
from pdf_reader.controllers.pdf_reader_controller import PDFController

pdf_controller = PDFController()

@require_GET
def open_pdf(request, file_path):
    decoded_file_name = unquote(file_path)  # Decode the file path from URL encoding
    try:
        success, message = pdf_controller.open_pdf(decoded_file_name)
        status = 200 if success else 500
    except Exception as e:
        print(f"Error opening PDF: {e}")
        return JsonResponse({"message": f"Failed to open PDF: {str(e)}"}, status=500)

    return JsonResponse({"message": message}, status=status)


@require_GET
def close_pdf(request):
    success, message = pdf_controller.close_pdf()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def scroll_up(request):
    success, message = pdf_controller.scroll_up()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def scroll_down(request):
    success, message = pdf_controller.scroll_down()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def zoom_in(request):
    success, message = pdf_controller.zoom_in()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def zoom_out(request):
    success, message = pdf_controller.zoom_out()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def enable_read_mode(request):
    success, message = pdf_controller.enable_read_mode()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def disable_read_mode(request):
    success, message = pdf_controller.disable_read_mode()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def save_pdf(request):
    success, message = pdf_controller.save_pdf()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def print_pdf(request):
    success, message = pdf_controller.print_pdf()
    status = 200 if success else 500
    return JsonResponse({"message": message}, status=status)

@require_GET
def pdf_controls(request):
    return JsonResponse({"message": "PDF controls loaded", "status": "success"})
