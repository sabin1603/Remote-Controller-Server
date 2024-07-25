import json
import subprocess

from django.http import JsonResponse
from django.views.decorators.http import require_http_methods, require_POST
from django.conf import settings
import os

def open_teams(request):
    try:
        teams_path = r"C:\Users\Sabin\AppData\Local\Microsoft\Teams\current\Teams.exe"
        subprocess.run([teams_path], check=True)  # Adjust this command as needed for your environment
        return JsonResponse({"message": "Microsoft Teams opened successfully."})
    except Exception as e:
        return JsonResponse({"message": f"Failed to open Microsoft Teams: {e}"}, status=500)


@require_http_methods(["GET"])
def get_contacts(request):
    # Path to the sample data file
    sample_file_path = os.path.join(settings.BASE_DIR, 'contacts_sample.json')

    try:
        # Load sample data from the file
        with open(sample_file_path, 'r') as file:
            contacts_data = json.load(file)

        return JsonResponse(contacts_data, safe=False)

    except FileNotFoundError:
        return JsonResponse({"error": "Sample data file not found"}, status=404)
    except json.JSONDecodeError:
        return JsonResponse({"error": "Error decoding sample data"}, status=500)

@require_POST
def start_call(request):
    try:
        # Extract user ID from request body
        data = json.loads(request.body)
        user_id = data.get('userId')

        if not user_id:
            return JsonResponse({"error": "User ID is required"}, status=400)

        # Construct the command to initiate a call using Teams
        teams_path = r"C:\Users\Sabin\AppData\Local\Microsoft\Teams\current\Teams.exe"
        # Example command; you may need to adjust it based on how Teams can be called
        call_command = [teams_path, "--call", user_id]

        # Run the command to initiate the call
        subprocess.run(call_command, check=True)

        return JsonResponse({"message": "Call initiated successfully."})
    except Exception as e:
        return JsonResponse({"message": f"Failed to initiate call: {e}"}, status=500)