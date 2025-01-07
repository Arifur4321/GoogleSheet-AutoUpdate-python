import os
import time
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
    except Exception as e:
        log(f"Error during Google Sheets authentication: {e}")
    return None

# Fetch Leads Matching the Email
def fetch_leads_by_email(email):
    try:
        response = api.get('lead/', params={"query": email})
        leads = response.get("data", [])
        log(f"{len(leads)} leads fetched for email {email}.")
        return [
            {"lead_id": lead["id"], "name": lead.get("display_name", "Unknown Lead")}
            for lead in leads
        ]
    except Exception as e:
        log(f"Error fetching leads by email: {e}")
        return []

# Fetch Call Data for the Last 3 Days
def fetch_calls(lead_id=None):
    today = datetime.date.today()
    last_3_days = [(today - datetime.timedelta(days=i)).strftime("%Y-%m-%d") for i in range(3)]
    all_data = []

    for date in last_3_days:
        params = {
            "date_start": date,
            "date_end": date,
            "lead_id": lead_id
        }

        try:
            response = api.get('activity/call', params=params)
            for record in response.get("data", []):
                record["date"] = date
                all_data.append(record)
        except Exception as e:
            log(f"Error fetching call data for date {date}: {e}")

    return all_data

# Process Data into Tabular Format
def process_close_data(data):
    if not data:
        log("Error: No call data found.")
        return []
    results = []
    for call in data:
        date = call.get('date', "Unknown Date")
        user = call.get('user_name', "Unknown User")
        lead_name = call.get('lead_name', "Unknown Lead")
        call_count = call.get('total_calls', 0)
        answered = call.get('answered_calls', 0)
        response_rate = round((answered / call_count) * 100, 2) if call_count else 0
        minutes = call.get('duration', 0)
        results.append([date, user, lead_name, call_count, answered, response_rate, minutes])
    log("Close API data processed successfully.")
    return results

# Write Data to Google Sheets
def write_to_google_sheet(sheet, data):
    if not sheet:
        log("Google Sheets service not authenticated.")
        return
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
    email = "riccardo@giacomofreddi.it"
    sheet_service = authenticate_google_sheets()
    if not sheet_service:
        log("Script terminated due to authentication failure.")
        return

    # Fetch Leads
    leads = fetch_leads_by_email(email)

    if not leads:
        log(f"No leads found for email {email}.")
        return

    # Fetch Calls for Leads
    call_data = []
    for lead in leads:
        log(f"Fetching calls for lead: {lead['name']}")
        call_data.extend(fetch_calls(lead_id=lead["lead_id"]))

    # Process and Write Data
    if not call_data:
        log(f"No call data found for email {email}.")
        return

    processed_data = process_close_data(call_data)
    header = [["Date", "User", "Lead Name", "Total Calls", "Answered Calls", "Response Rate (%)", "Duration (seconds)"]]
    full_data = header + processed_data
    write_to_google_sheet(sheet_service, full_data)
    log("Script finished successfully.")

if __name__ == "__main__":
    main()
