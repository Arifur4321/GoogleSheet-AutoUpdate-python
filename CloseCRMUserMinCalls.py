import os
import time
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

# Fetch all active team members
def fetch_active_team_members():
    team_members = []
    offset = 0
    limit = 100

    while True:
        try:
            response = api.get("user/", params={"_skip": offset, "_limit": limit})
            members = response.get("data", [])
            active_members = [member for member in members if member.get("enabled", False)]  # Only active users
            team_members.extend(active_members)
            offset += limit
            if not response.get("has_more", False):
                break
        except Exception as e:
            log(f"Error fetching team members: {e}")
            break

    log(f"Total active team members fetched: {len(team_members)}.")
    return team_members

# Fetch all calls by date
def fetch_calls_by_date(date):
    all_calls = []
    offset = 0
    limit = 100
    while True:
        try:
            params = {"date_start": date, "date_end": date, "_skip": offset, "_limit": limit}
            response = api.get("activity/call", params=params)
            calls = response.get("data", [])
            all_calls.extend(calls)
            offset += limit
            if not response.get("has_more", False):
                break
        except Exception as e:
            log(f"Error fetching calls for date {date} (offset {offset}): {e}")
            break
    log(f"Fetched {len(all_calls)} calls for date: {date}.")
    return all_calls

# Process data for Google Sheets
def process_data_for_sheet(team_members, dates):
    header = ["NAME", "COGNOME/SURNAME", "EMAIL"] + dates
    rows = []

    # Fetch all calls for the specified dates
    all_calls = {}
    for date in dates:
        all_calls[date] = fetch_calls_by_date(date)

    # Process data for each team member
    for member in team_members:
        row = [member["first_name"], member["last_name"], member["email"]]
        for date in dates:
            user_calls = [call for call in all_calls[date] if call.get("user_id") == member["id"]]
            call_count = len(user_calls)
            total_minutes = round(sum(call.get("duration", 0) / 60 for call in user_calls), 2)
            row.append(f"Calls: {call_count}, Min: {total_minutes}")
        rows.append(row)

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

    active_team_members = fetch_active_team_members()
    if not active_team_members:
        log("No active team members found in Close CRM.")
        return

    # Define the date range for the last 5 days
    today = datetime.date.today()
    dates = [(today - datetime.timedelta(days=i)).strftime("%d/%m/%Y") for i in range(5, 0, -1)]

    data = process_data_for_sheet(active_team_members, dates)
    write_to_sheet(sheet_service, data)
    log("Script finished successfully.")

if __name__ == "__main__":
    main()
