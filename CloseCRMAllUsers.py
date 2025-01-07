import os
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from closeio_api import Client
import datetime

# Constants
SHEET_ID = "1-1rE5epVoYSJgqgkNYsGXuMrYQWkjbPjYpQUYgzMOTI"
RANGE_NAME = "Close Data!A1"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'F:\\Giacometti\\New-Working-Directory-25-10-2024\\Python-report\\SERVICE_ACCOUNT_FILE\\closecrmdata-7f2f781a56e5.json'
CLOSE_API_KEY = "api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i"

# Initialize Close API Client

api = Client(CLOSE_API_KEY) 

# Log helper
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

# Fetch all team members from Close CRM (with pagination)
def fetch_all_team_members():
    all_users = []
    has_more = True
    offset = 0
    limit = 100  # Close API default pagination limit

    while has_more:
        try:
            response = api.get('user/', params={'_skip': offset, '_limit': limit})
            users = response.get('data', [])
            all_users.extend(users)
            has_more = response.get('has_more', False)
            offset += limit
            log(f"Fetched {len(users)} team members (offset: {offset}).")
        except Exception as e:
            log(f"Error fetching team members: {e}")
            has_more = False

    log(f"Total team members fetched: {len(all_users)}.")
    return [
        {
            "id": user.get("id"),
            "first_name": user.get("first_name", "N/A"),
            "last_name": user.get("last_name", "N/A"),
            "email": user.get("email", "N/A"),
            "is_active": user.get("is_active", False),  # Explicitly get 'is_active'
            "created": user.get("created", "N/A"),
            "updated": user.get("updated", "N/A")
        }
        for user in all_users
    ]

# Process all team members into tabular format for Google Sheets
def process_team_data(team_members):
    header = ["ID", "First Name", "Last Name", "Email", "Active Status", "Created Date", "Last Updated"]
    rows = [[
        member["id"],
        member["first_name"],
        member["last_name"],
        member["email"],
        "Active" if member["is_active"] else "Inactive",  # Correctly interpret 'is_active'
        member["created"],
        member["updated"]
    ] for member in team_members]
    return [header] + rows

# Write data to Google Sheets
def write_to_sheet(sheet, data):
    try:
        body = {"values": data}
        result = sheet.values().update(
            spreadsheetId=SHEET_ID, range=RANGE_NAME,
            valueInputOption="RAW", body=body).execute()
        log(f"{result.get('updatedCells')} cells updated in Google Sheets.")
    except Exception as e:
        log(f"Error writing to Google Sheets: {e}")

# Main Script
def main():
    log("Script started.")
    sheet_service = authenticate_google_sheets()
    if not sheet_service:
        log("Script terminated due to authentication failure.")
        return

    all_users = fetch_all_team_members()
    if not all_users:
        log("No users found in Close CRM.")
        return

    # Debug: Log active/inactive counts
    active_users = sum(1 for user in all_users if user["is_active"])
    inactive_users = len(all_users) - active_users
    log(f"Active users: {active_users}, Inactive users: {inactive_users}")

    processed_data = process_team_data(all_users)
    write_to_sheet(sheet_service, processed_data)
    log("Script finished successfully.")

if __name__ == "__main__":
    main()
