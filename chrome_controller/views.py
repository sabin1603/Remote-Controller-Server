from django.http import JsonResponse
from django.views.decorators.http import require_POST, require_GET
from django.views.decorators.csrf import csrf_exempt
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import os

# Setup Selenium WebDriver
chrome_service = Service('D:\ChromeDriver\chromedriver-win64\chromedriver.exe')  # Specify the path to your ChromeDriver
chrome_options = Options()
#chrome_options.add_argument("--remote-debugging-port=9222")  # Optional: for debugging
driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

@require_GET
def open_chrome(request):
    try:
        driver.get('http://www.google.com')  # Open Chrome and navigate to a URL
        return JsonResponse({"message": "Chrome opened successfully."})
    except Exception as e:
        return JsonResponse({"message": f"Error opening Chrome: {e}"}, status=500)

@csrf_exempt
@require_POST
def new_tab(request):
    try:
        driver.execute_script("window.open('');")  # Open a new tab
        driver.switch_to.window(driver.window_handles[-1])  # Switch to the new tab
        return JsonResponse({"message": "New tab opened"})
    except Exception as e:
        return JsonResponse({"message": f"Error opening new tab: {e}"}, status=500)

@csrf_exempt
@require_POST
def go_home(request):
    try:
        driver.get('http://www.google.com')  # Navigate to the home page
        return JsonResponse({"message": "Navigated to home"})
    except Exception as e:
        return JsonResponse({"message": f"Error navigating to home: {e}"}, status=500)

@csrf_exempt
@require_POST
def go_back(request):
    try:
        driver.execute_script("window.history.back();")  # Go back in history
        return JsonResponse({"message": "Navigated back"})
    except Exception as e:
        return JsonResponse({"message": f"Error navigating back: {e}"}, status=500)

@csrf_exempt
@require_POST
def go_forward(request):
    try:
        driver.execute_script("window.history.forward();")  # Go forward in history
        return JsonResponse({"message": "Navigated forward"})
    except Exception as e:
        return JsonResponse({"message": f"Error navigating forward: {e}"}, status=500)

@csrf_exempt
@require_POST
def close_current_tab(request):
    try:
        driver.close()  # Close the current tab
        driver.switch_to.window(driver.window_handles[0])  # Switch to the remaining tab
        return JsonResponse({"message": "Current tab closed"})
    except Exception as e:
        return JsonResponse({"message": f"Error closing current tab: {e}"}, status=500)

@csrf_exempt
@require_POST
def zoom_in(request):
    try:
        driver.execute_script("document.body.style.zoom = '125%'")  # Zoom in
        return JsonResponse({"message": "Zoomed in"})
    except Exception as e:
        return JsonResponse({"message": f"Error zooming in: {e}"}, status=500)

@csrf_exempt
@require_POST
def zoom_out(request):
    try:
        driver.execute_script("document.body.style.zoom = '75%'")  # Zoom out
        return JsonResponse({"message": "Zoomed out"})
    except Exception as e:
        return JsonResponse({"message": f"Error zooming out: {e}"}, status=500)

@csrf_exempt
@require_POST
def scroll_up(request):
    try:
        driver.execute_script("window.scrollBy(0, -100)")  # Scroll up
        return JsonResponse({"message": "Scrolled up"})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling up: {e}"}, status=500)

@csrf_exempt
@require_POST
def scroll_down(request):
    try:
        driver.execute_script("window.scrollBy(0, 100)")  # Scroll down
        return JsonResponse({"message": "Scrolled down"})
    except Exception as e:
        return JsonResponse({"message": f"Error scrolling down: {e}"}, status=500)

@csrf_exempt
@require_POST
def go_to_left_tab(request):
    try:
        driver.execute_script("window.history.go(-1)")  # Go to the previous tab
        return JsonResponse({"message": "Navigated to left tab"})
    except Exception as e:
        return JsonResponse({"message": f"Error navigating to left tab: {e}"}, status=500)

@csrf_exempt
@require_POST
def go_to_right_tab(request):
    try:
        driver.execute_script("window.history.go(1)")  # Go to the next tab
        return JsonResponse({"message": "Navigated to right tab"})
    except Exception as e:
        return JsonResponse({"message": f"Error navigating to right tab: {e}"}, status=500)

@csrf_exempt
@require_POST
def close_chrome(request):
    try:
        driver.quit()  # Close Chrome
        return JsonResponse({"message": "Chrome closed"})
    except Exception as e:
        return JsonResponse({"message": f"Error closing Chrome: {e}"}, status=500)

@csrf_exempt
@require_POST
def navigate(request):
    try:
        data = json.loads(request.body)
        query = data.get('query')
        driver.get(query)
        return JsonResponse({"message": "Navigation successful"})
    except Exception as e:
        return JsonResponse({"message": f"Error navigating: {e}"}, status=500)

@csrf_exempt
@require_POST
def navigate_up(request):
    try:
        driver.execute_script("window.scrollBy(0, -100)")  # Scroll up
        return JsonResponse({"message": "Navigated up"})
    except Exception as e:
        return JsonResponse({"message": f"Error navigating up: {e}"}, status=500)

@csrf_exempt
@require_POST
def navigate_down(request):
    try:
        driver.execute_script("window.scrollBy(0, 100)")  # Scroll down
        return JsonResponse({"message": "Navigated down"})
    except Exception as e:
        return JsonResponse({"message": f"Error navigating down: {e}"}, status=500)
