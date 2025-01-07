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

# Find email in Close API
def find_email_in_close(email):
    log(f"Searching for email: {email.strip().lower()}")
    email = email.strip().lower()
    user_id = None
    lead_id = None

    # Search in users
    try:
        users = api.get("me/").get("data", [])
        for user in users:
            if user.get("email", "").strip().lower() == email:
                user_id = user.get("id")
                log(f"Email found in users: {email}")
                return user_id, lead_id
        log("Email not found in users.")
        
    except Exception as e:
        log(f"Error searching for email in users: {e}")

    # Search in leads
    try:
        leads = api.get("lead/", params={"query": f'"{email}"'}).get("data", [])
        for lead in leads:
            for contact in lead.get("contacts", []):
                if email in [contact_email.get("email", "").strip().lower() for contact_email in contact.get("emails", [])]:
                    lead_id = lead.get("id")
                    log(f"Email found in leads: {email}")
                    return user_id, lead_id
        log("Email not found in leads.")
    except Exception as e:
        log(f"Error searching for email in leads: {e}")

    return user_id, lead_id

# Fetch activity metrics
def fetch_metrics(user_id=None, lead_id=None, start_date=None, end_date=None):
    metrics = {
        "created_leads": 0,
        "outbound_calls": 0,
        "inbound_calls": 0,
        "avg_call_duration": 0,
        "sent_emails": 0,
        "received_emails": 0,
        "opportunities_created": 0,
        "opportunities_won": 0,
    }
    try:
        # Fetch call activities
        params = {
            "date_start": start_date,
            "date_end": end_date,
            "user_id": user_id,
            "lead_id": lead_id
        }
        params = {k: v for k, v in params.items() if v}
        calls = api.get("activity/call", params=params).get("data", [])
        for call in calls:
            if call.get("direction") == "outbound":
                metrics["outbound_calls"] += 1
            elif call.get("direction") == "inbound":
                metrics["inbound_calls"] += 1
            metrics["avg_call_duration"] += call.get("duration", 0)

        # Fetch emails
        emails = api.get("activity/email", params=params).get("data", [])
        for email in emails:
            if email.get("status") == "sent":
                metrics["sent_emails"] += 1
            elif email.get("status") == "received":
                metrics["received_emails"] += 1

        # Fetch leads
        leads = api.get("lead/", params={"date_start": start_date, "date_end": end_date}).get("data", [])
        metrics["created_leads"] = len(leads)

        # Fetch opportunities (stub example as Close API might need additional params)
        # Assuming opportunities are related to leads
        metrics["opportunities_created"] = 0
        metrics["opportunities_won"] = 0

        metrics["avg_call_duration"] = round(metrics["avg_call_duration"] / (metrics["outbound_calls"] + metrics["inbound_calls"]) if (metrics["outbound_calls"] + metrics["inbound_calls"]) > 0 else 0, 2)
        return metrics
    except Exception as e:
        log(f"Error fetching metrics: {e}")
        return metrics

# Process data into tabular format
def process_metrics(metrics):
    return [
        ["Metric", "Value"],
        ["Created Leads", metrics["created_leads"]],
        ["Outbound Calls", metrics["outbound_calls"]],
        ["Inbound Calls", metrics["inbound_calls"]],
        ["Average Call Duration (minutes)", metrics["avg_call_duration"]],
        ["Sent Emails", metrics["sent_emails"]],
        ["Received Emails", metrics["received_emails"]],
        ["Opportunities Created", metrics["opportunities_created"]],
        ["Opportunities Won", metrics["opportunities_won"]],
    ]

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

    email = "samuele@giacomofreddi.it"
    start_date = "2024-12-03"
    end_date = "2024-12-05"

    user_id, lead_id = find_email_in_close(email)
    if not user_id and not lead_id:
        log(f"No user or lead found with email: {email}")
        return

    metrics = fetch_metrics(user_id=user_id, lead_id=lead_id, start_date=start_date, end_date=end_date)
    processed_data = process_metrics(metrics)
    write_to_sheet(sheet_service, processed_data)
    log("Script finished successfully.")

if __name__ == "__main__":
    main()
