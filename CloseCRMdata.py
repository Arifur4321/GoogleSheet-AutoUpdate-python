import os
import time
import random
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from closeio_api import Client
import datetime

# Google Sheets Setup
SHEET_ID = "1-1rE5epVoYSJgqgkNYsGXuMrYQWkjbPjYpQUYgzMOTI"  # Replace with your Google Sheet ID
RANGE_NAME = "Close Data!A1"  # Adjust to your sheet's range
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'F:\\Giacometti\\New-Working-Directory-25-10-2024\\Python-report\\SERVICE_ACCOUNT_FILE\\closecrmdata-7f2f781a56e5.json'  # Correct JSON file path

# Close API Setup
CLOSE_API_KEY = "api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i"

# Initialize Close API Client
api = Client(CLOSE_API_KEY)

#email : lorenzobanchetti@giacomofreddi.it

# Log helper
def log(message):
    print(f"[{datetime.datetime.now()}] {message}")

# Authenticate with Google Sheets
def authenticate_google_sheets():
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        log("Error: Service account file not found. Check the path.")
        return None
    try:
        credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials)
        log("Google Sheets authenticated successfully.")
        return service.spreadsheets()
    except FileNotFoundError as e:
        log(f"Error: Service account file not found. Ensure the file exists at {SERVICE_ACCOUNT_FILE}.")
    except Exception as e:
        log(f"Error during Google Sheets authentication: {e}")
    return None

# Fetch Team Members
def fetch_team_members():
    try:
        response = api.get('user/')
        log("Team members fetched successfully.")
        users = response.get("data", [])
        return [
            {
                "name": f"{user.get('first_name', 'Unknown')} {user.get('last_name', 'Unknown')}".strip(),
                "user_id": user["id"]
            }
            for user in users
        ]
    except Exception as e:
        log(f"Error fetching team members: {e}")
        return []

# Fetch Data for All Team Members from Close API for Last 2 Days
def fetch_close_data(team_members):
    all_data = []
    today = datetime.date.today()
    last_2_days = [(today - datetime.timedelta(days=i)).strftime("%Y-%m-%d") for i in range(2)]

    for date in last_2_days:
        for member in team_members:
            params = {
                "date_start": date,  # Fetch data for a specific date
                "date_end": date,
                "user_id": member['user_id']  # Filter by user_id
            }
            try:
                response = api.get('activity/call', params=params)
                log(f"Data fetched successfully for {member['name']} on {date}.")
                for record in response.get("data", []):
                    record["user_name"] = member["name"]
                    record["date"] = date
                    all_data.append(record)
            except Exception as e:
                log(f"Error fetching data for {member['name']} on {date}: {e}")
    return {"data": all_data}

# Process Data into Tabular Format example code
def process_close_data(data):
    if not data or 'data' not in data:
        log("Error: Invalid data format received from Close API.")
        return []
    results = []
    for call in data['data']:
        date = call.get('date', "Unknown Date")
        user = call.get('user_name', "Unknown User")
        call_count = call.get('total_calls', 0)
        answered = call.get('answered_calls', 0)
        response_rate = round((answered / call_count) * 100, 2) if call_count else 0
        minutes = call.get('total_minutes', 0)
        results.append([date, user, call_count, answered, response_rate, minutes])
    log("Close API data processed successfully.")
    return results

# Write Data to Google Sheets with Retry Mechanism
def write_with_retries(sheet, data, retries=3, delay=10):
    for attempt in range(retries):
        try:
            body = {"values": data}
            result = sheet.values().update(
                spreadsheetId=SHEET_ID, range=RANGE_NAME,
                valueInputOption="RAW", body=body).execute()
            log(f"{result.get('updatedCells')} cells updated in Google Sheets.")
            return  # Exit the function on success
        except Exception as e:
            log(f"Error writing to Google Sheets (Attempt {attempt + 1}): {e}")
            if "RATE_LIMIT_EXCEEDED" in str(e):
                log(f"Retrying in {delay} seconds...")
                time.sleep(delay)
            else:
                break  # Exit on non-rate-limit errors
    log("Failed to update Google Sheets after multiple retries.")

# Main Script
def main():
    log("Script started.")
    sheet_service = authenticate_google_sheets()
    if not sheet_service:
        log("Script terminated due to authentication failure.")
        return

    team_members = fetch_team_members()
    if not team_members:
        log("Script terminated due to team member fetch failure.")
        return

    close_data = fetch_close_data(team_members)
    if not close_data:
        log("Script terminated due to Close API data fetch failure.")
        return

    processed_data = process_close_data(close_data)
    header = [["Date", "User", "Total Calls", "Answered Calls", "Response Rate (%)", "Minutes in Call"]]
    full_data = header + processed_data

    # Write data to Google Sheets with retries
    write_with_retries(sheet_service, full_data)
    log("Script finished successfully.")

if __name__ == "__main__":
    main()
