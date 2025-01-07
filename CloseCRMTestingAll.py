import os
import datetime
import random
import time
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from closeio_api import Client

# Constants
SHEET_ID = "1-1rE5epVoYSJgqgkNYsGXuMrYQWkjbPjYpQUYgzMOTI"
RANGE_NAME = "Close Data!A1"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'F:\\Giacometti\\New-Working-Directory-25-10-2024\\Python-report\\SERVICE_ACCOUNT_FILE\\closecrmdata-7f2f781a56e5.json'
CLOSE_API_KEY = "api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i"

# Initialize Close API Client
api = Client(CLOSE_API_KEY)

# Log Helper
def log(message):
    print(f"[{datetime.datetime.now()}] {message}")

# Authenticate with Google Sheets
def authenticate_google_sheets():
    try:
        credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials)
        log("Google Sheets authenticated successfully.")
        return service.spreadsheets()
    except Exception as e:
        log(f"Error during Google Sheets authentication: {e}")
        return None

# Fetch All Salespersons
def fetch_salespersons():
    try:
        response = api.get('user/')
        log("Salespersons fetched successfully.")
        users = response.get("data", [])
        return [
            {"name": f"{user.get('first_name', 'Unknown')} {user.get('last_name', 'Unknown')}".strip(), "user_id": user["id"]}
            for user in users
        ]
    except Exception as e:
        log(f"Error fetching salespersons: {e}")
        return []

# Fetch Call and Email Activities for Last 3 Days
def fetch_activities(user_id):
    all_data = []
    today = datetime.date.today()
    for i in range(3):
        date = (today - datetime.timedelta(days=i)).strftime("%Y-%m-%d")
        try:
            # Fetch Calls
            call_response = api.get('activity/call', params={"date_start": date, "date_end": date, "user_id": user_id})
            for call in call_response.get("data", []):
                call["activity_type"] = "Call"
                all_data.append(call)

            # Fetch Emails
            email_response = api.get('activity/email', params={"date_start": date, "date_end": date, "user_id": user_id})
            for email in email_response.get("data", []):
                email["activity_type"] = "Email"
                all_data.append(email)
            
            log(f"Activities fetched for user_id {user_id} on {date}.")
        except Exception as e:
            log(f"Error fetching activities for user_id {user_id} on {date}: {e}")
    return all_data

# Process Activities into a Tabular Format
def process_activities(activities):
    if not activities:
        log("No activities found.")
        return []

    results = []
    for activity in activities:
        date = activity.get('date', "Unknown Date")[:10]
        activity_type = activity.get('activity_type', "Unknown Type")
        duration = activity.get('duration', 0) / 60 if activity_type == "Call" else "N/A"
        direction = activity.get('direction', "N/A")
        lead_name = activity.get('lead_name', "Unknown Lead")
        user_name = activity.get('user_name', "Unknown User")
        email_subject = activity.get('subject', "N/A") if activity_type == "Email" else "N/A"
        results.append([date, activity_type, user_name, lead_name, duration, direction, email_subject])
    log("Activities processed successfully.")
    return results

# Write Data to Google Sheets
def write_to_google_sheets(sheet, data):
    if not sheet:
        log("Google Sheets service not authenticated.")
        return
    try:
        # Define headers
        headers = [["Date", "Activity Type", "User", "Lead/Client", "Duration (min)", "Direction", "Email Subject"]]
        body = {"values": headers + data}
        result = sheet.values().update(
            spreadsheetId=SHEET_ID, range=RANGE_NAME,
            valueInputOption="RAW", body=body).execute()
        log(f"{result.get('updatedCells')} cells updated in Google Sheets.")
    except Exception as e:
        log(f"Error writing to Google Sheets: {e}")

# Main Script
def main():
    log("Script started.")
    
    # Authenticate with Google Sheets
    sheet_service = authenticate_google_sheets()
    if not sheet_service:
        log("Script terminated due to Google Sheets authentication failure.")
        return

    # Fetch salespersons
    salespersons = fetch_salespersons()
    if not salespersons:
        log("Script terminated due to salesperson fetch failure.")
        return

    all_activities = []
    for salesperson in salespersons:
        user_id = salesperson["user_id"]
        user_activities = fetch_activities(user_id)
        all_activities.extend(user_activities)

    # Process data
    processed_data = process_activities(all_activities)

    # Write to Google Sheets
    write_to_google_sheets(sheet_service, processed_data)
    log("Script completed successfully.")

if __name__ == "__main__":
    main()
