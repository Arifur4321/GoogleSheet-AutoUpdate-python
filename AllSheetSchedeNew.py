# Import necessary libraries
from closeio_api import Client
from googleapiclient.discovery import build
from google.oauth2 import service_account
import logging
from datetime import datetime, timedelta

# Set up logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Constants for Close CRM and Google Sheets credentials
SHEET_ID = "1-1rE5epVoYSJgqgkNYsGXuMrYQWkjbPjYpQUYgzMOTI"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'F:\\Giacometti\\New-Working-Directory-25-10-2024\\Python-report\\SERVICE_ACCOUNT_FILE\\closecrmdata-7f2f781a56e5.json'

CLOSE_API_KEY = "api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i"

# Initialize Close API client
api = Client(CLOSE_API_KEY)

# Function to fetch all team members
def fetch_all_team_members():
    all_users = []
    has_more = True
    offset = 0
    limit = 100

    while has_more:
        try:
            response = api.get('user/', params={'_skip': offset, '_limit': limit})
            users = response.get('data', [])
            all_users.extend(users)
            has_more = response.get('has_more', False)
            offset += limit
        except Exception as e:
            logging.error(f"Error fetching team members: {e}")
            has_more = False

    return [
        {
            "id": user.get("id"),
            "email": user.get("email", "N/A"),
        }
        for user in all_users
    ]

# Function to get user IDs by email
def get_user_ids_by_email(email):
    all_users = fetch_all_team_members()
    return [user['id'] for user in all_users if user['email'] == email]

# Function to fetch Close CRM data

def fetch_close_data(email, start_date, end_date):
    user_ids = get_user_ids_by_email(email)
    if not user_ids:
        logging.warning(f"No user IDs found for email {email}. Skipping.")
        return 0, 0

    try:
        data = {
            "datetime_range": {
                "start": f"{start_date}T00:00:00Z",
                "end": f"{end_date}T23:59:59Z",
            },
            "users": user_ids,
            "type": "overview",
            "metrics": [
                "calls.outbound.all.count",
                "calls.outbound.all.sum_duration"
            ],
        }
        response = api.post('report/activity', data=data)
        metrics = response.get("aggregations", {}).get("totals", {})
        total_calls = metrics.get("calls.outbound.all.count", 0)
        total_duration = metrics.get("calls.outbound.all.sum_duration", 0)
        return total_calls, total_duration
    except Exception as e:
        logging.error(f"Error fetching Close CRM data for {email}: {e}")
        return 0, 0

# Function to format duration in seconds new 
def format_duration(seconds):
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    if hours > 0:
        return f"{hours}h{minutes}min" if seconds == 0 else f"{hours}h{minutes}min{seconds}sec"
    elif minutes > 0:
        return f"{minutes}min{seconds}sec" if seconds > 0 else f"{minutes}min"
    else:
        return f"{seconds}sec"

# Function to fetch sheet names from the "Schede" column in the "Settings" sheet
def fetch_sheet_names_from_settings():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SHEET_ID, range="Settings!B2:B").execute()
    sheet_names = [row[0] for row in result.get('values', []) if row]
    logging.info(f"All sheet names from the settings: {sheet_names}")
    return sheet_names

# Function to fetch Google Sheet data for a specific sheet
def fetch_google_sheet_data(sheet_name):
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SHEET_ID, range=f"{sheet_name}!A1:ZZZ").execute()  # Fetch all columns dynamically
    data = result.get('values', [])
    return data, service, sheet

# Function to update Google Sheet for a specific sheet
def update_google_sheet(service, sheet, cell_range, value):
    body = {'values': [[value]]}
    try:
        sheet.values().update(spreadsheetId=SHEET_ID, range=cell_range, valueInputOption="RAW", body=body).execute()
    except Exception as e:
        logging.error(f"Error updating Google Sheet at {cell_range}: {e}")

# Function to calculate column name
def get_column_letter(index):
    """Convert a zero-based column index to a column letter (e.g., 0 -> A, 701 -> ZZ)."""
    column = ""
    while index >= 0:
        column = chr(index % 26 + 65) + column
        index = index // 26 - 1
    return column

# Function to validate dates in the header row
def validate_dates_in_header(headers, expected_start_date, expected_end_date):
    """Validate that all dates between start_date and end_date are in the header."""
    date_range = [
        (expected_start_date + timedelta(days=i)).strftime('%d/%m/%Y')
        for i in range((expected_end_date - expected_start_date).days + 1)
    ]
    missing_dates = [date for date in date_range if date not in headers]
    if missing_dates:
        logging.warning(f"Missing dates in header: {missing_dates}")
    else:
        logging.info("All dates are present in the header.")

# Function to process and update a sheet
def process_and_update_sheet(sheet_name):
    logging.info(f"Processing sheet: {sheet_name}")
    sheet_data, service, sheet = fetch_google_sheet_data(sheet_name)
    if not sheet_data:
        logging.warning(f"Sheet {sheet_name} is empty or unavailable.")
        return

    headers = sheet_data[0]  # Extract header row (e.g., dates)
    logging.info(f"Header row for sheet {sheet_name}: {headers}")

    # Validate all dates in the header
    expected_start_date = datetime.strptime('01/11/2024', '%d/%m/%Y')
    expected_end_date = datetime.strptime('30/11/2024', '%d/%m/%Y')
    validate_dates_in_header(headers, expected_start_date, expected_end_date)

    for row_index, row in enumerate(sheet_data[1:], start=2):  # Skip header row
        if len(row) > 2 and row[2]:  # Identify rows with email addresses
            email = row[2]
            logging.info(f"Processing data for email: {email} in sheet: {sheet_name}")

            call_row = None
            duration_row = None
            for i in range(row_index, len(sheet_data)):
                if len(sheet_data[i]) > 0:
                    if sheet_data[i][0].strip().lower() == "call":  # Locate "Call" row
                        call_row = i + 1
                    if sheet_data[i][0].strip().lower() == "duration":  # Locate "Duration" row
                        duration_row = i + 1
                if call_row and duration_row:
                    break

            if not call_row or not duration_row:
                logging.warning(f"Could not locate 'Call' or 'Duration' rows for email: {email} in sheet: {sheet_name}")
                continue

            for col_index, date in enumerate(headers[3:], start=3):  # Dates start from column D
                column_letter = get_column_letter(col_index)
                logging.info(f"Processing date {date} at column {column_letter}")
                if date:
                    try:
                        formatted_date = datetime.strptime(date, '%d/%m/%Y').strftime('%Y-%m-%d')
                        total_calls, total_duration = fetch_close_data(email, formatted_date, formatted_date)

                        # Update the "Call" row
                        call_cell_range = f"{sheet_name}!{column_letter}{call_row}"
                        update_google_sheet(service, sheet, call_cell_range, total_calls)

                        # Update the "Duration" row
                        formatted_duration = format_duration(total_duration)
                        duration_cell_range = f"{sheet_name}!{column_letter}{duration_row}"
                        update_google_sheet(service, sheet, duration_cell_range, formatted_duration)

                        logging.info(f"Updated {call_cell_range} with {total_calls}")
                        logging.info(f"Updated {duration_cell_range} with {formatted_duration}")

                    except Exception as e:
                        logging.error(f"Error processing date {date} for email {email} in sheet: {sheet_name}: {e}")
                        continue

# Main function to process sheets listed in the "Settings" sheet
def main():
    sheet_names = fetch_sheet_names_from_settings()
    for sheet_name in sheet_names:
        logging.info(f"Starting processing for sheet: {sheet_name}")
        process_and_update_sheet(sheet_name)
    logging.info("Processing completed for all sheets.")

if __name__ == "__main__":
    main()
