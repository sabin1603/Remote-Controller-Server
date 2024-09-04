import os
import subprocess
import psutil
import skpy
from django.http import JsonResponse
from django.views.decorators.http import require_GET
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
from skpy import Skype

# Global Skype process instance
current_skype_process = None

EMAIL = os.getenv('SKYPE_EMAIL')
PASSWORD = os.getenv('SKYPE_PASSWORD')


@require_GET
def get_skype_contacts(request):
    """Returns a list of Skype contacts."""
    try:
        # Authenticate with Skype using email
        sk = Skype(EMAIL, PASSWORD)

        # Fetch contacts
        contacts = sk.contacts
        contact_list = [
            {'username': contact.id, 'display_name': str(contact.name)}  # Ensure name is a string
            for contact in contacts
        ]

        return JsonResponse({'contacts': contact_list})
    except Exception as e:
        return JsonResponse({'message': f"Error fetching contacts: {e}"}, status=500)

@require_GET
def open_skype(request):
    """Open Skype application."""
    global current_skype_process

    # Check if Skype is already running
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == "Skype.exe":
            return JsonResponse({'status': 'success', 'message': 'Skype is already open.'})

    try:
        # Start Skype based on the OS
        if os.name == 'nt':  # Windows
            current_skype_process = subprocess.Popen(["start", "skype:"], shell=True)
        else:  # MacOS or Linux
            current_skype_process = subprocess.Popen(["open", "skype:"])

        return JsonResponse({'status': 'success', 'message': 'Skype opened successfully.'})
    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})


@require_GET
def close_skype(request):
    """
    Closes the Skype application.
    """
    global current_skype_process

    # Check if Skype is running using psutil, as the process handle might be lost
    skype_running = any(proc.info['name'] == "Skype.exe" for proc in psutil.process_iter(['pid', 'name']))

    if not skype_running:
        return JsonResponse({"message": "Skype is not currently open."}, status=400)

    try:
        # Attempt to close the Skype application using taskkill
        subprocess.run(["taskkill", "/IM", "Skype.exe", "/F"], check=True)

        # Ensure all Skype processes are killed
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == "Skype.exe":
                proc.kill()

        current_skype_process = None  # Clear the global process reference
        return JsonResponse({"message": "Skype closed successfully."})
    except Exception as e:
        return JsonResponse({"message": f"Error closing Skype: {e}"}, status=500)


@require_GET
def call_skype_user(request, username):
    """
    Initiates a call to a specific Skype user.
    """
    try:
        skype_uri = f"skype:{username}?call"
        if os.name == 'nt':  # Windows
            subprocess.Popen(["start", skype_uri], shell=True)
        else:  # MacOS or Linux
            subprocess.Popen(["open", skype_uri])
        return JsonResponse({"message": f"Calling {username} on Skype."})
    except Exception as e:
        return JsonResponse({"message": f"Error calling {username}: {e}"}, status=500)


@csrf_exempt
@require_POST
def send_message(request):
    """
    Sends a message to a specific Skype user using the skpy library.
    """
    try:
        username = request.POST.get('username')
        message = request.POST.get('message')

        if not username or not message:
            return JsonResponse({"message": "Username and message are required."}, status=400)

        # Initialize Skype session
        sk = Skype(EMAIL, PASSWORD)

        # Print username and message for debugging
        print(f"Attempting to send message to {username}")
        print(f"Message content: {message}")

        # Fetch the contact
        contact = sk.contacts[username]
        if not contact:
            return JsonResponse({"message": f"Contact not found for {username}"}, status=404)

        # Retrieve or create chat with the contact
        chat = contact.chat
        if chat:
            print("Chat object successfully retrieved.")
        else:
            print("Failed to retrieve chat object.")
            return JsonResponse({"message": f"Chat object is None for {username}"}, status=404)

        # Send the message
        chat.sendMsg(message)
        print("Message sent successfully.")
        return JsonResponse({"message": f"Message sent to {username} on Skype."})
    except Exception as e:
        print(f"Exception occurred: {e}")
        return JsonResponse({"message": f"Error sending message to {username}: {e}"}, status=500)

