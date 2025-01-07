from closeio_api import Client
from googleapiclient.discovery import build
from google.oauth2 import service_account
import logging
from datetime import datetime

# Set up logging works fine for one sheet 

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Constants credentials of Close CRM and google sheet 
SHEET_ID = "1-1rE5epVoYSJgqgkNYsGXuMrYQWkjbPjYpQUYgzMOTI"
RANGE_NAME = "Close Data!A1:Z"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'F:\\Giacometti\\New-Working-Directory-25-10-2024\\Python-report\\SERVICE_ACCOUNT_FILE\\closecrmdata-7f2f781a56e5.json'
# SERVICE_ACCOUNT_FILE = '/home/u121027207/public_html/CloseData/closecrmdata-7f2f781a56e5.json'

CLOSE_API_KEY = "api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i"

# Initialize Close API client
api = Client(CLOSE_API_KEY)

def fetch_all_team_members():
    """Fetch all team members from Close CRM."""
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
            logging.info(f"Fetched {len(users)} team members (offset: {offset}).")
        except Exception as e:
            logging.error(f"Error fetching team members: {e}")
            has_more = False

    logging.info(f"Total team members fetched: {len(all_users)}.")
    return [
        {
            "id": user.get("id"),
            "email": user.get("email", "N/A"),
        }
        for user in all_users
    ]

def get_user_ids_by_email(email):
    """Retrieve user ID(s) associated with the given email."""
    all_users = fetch_all_team_members()
    user_ids = [user['id'] for user in all_users if user['email'] == email]
    if not user_ids:
        logging.warning(f"No users found with email: {email}")
    return user_ids

def fetch_close_data(email, start_date, end_date):
    """Fetch Close CRM call data for a specific email and date range."""
    logging.info(f"Fetching Close CRM data for email: {email}, Date Range: {start_date} - {end_date}")
    user_ids = get_user_ids_by_email(email)
    if not user_ids:
        logging.error(f"No user IDs found for email {email}. Skipping.")
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
        logging.debug(f"Close API Response: {response}")

        metrics = response.get("aggregations", {}).get("totals", {})
        total_calls = metrics.get("calls.outbound.all.count", 0)
        total_duration = metrics.get("calls.outbound.all.sum_duration", 0)

        logging.info(f"Close CRM data fetched: {total_calls} Calls, {total_duration} sec Duration")
        return total_calls, total_duration
    except Exception as e:
        logging.error(f"Error fetching Close CRM data for {email}. Error: {e}")
        return 0, 0


def format_duration(seconds):
    """Format duration in seconds to human-readable format."""
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    if hours > 0:
        return f"{hours}h{minutes}min" if seconds == 0 else f"{hours}h{minutes}min{seconds}sec"
    elif minutes > 0:
        return f"{minutes}min{seconds}sec" if seconds > 0 else f"{minutes}min"
    else:
        return f"{seconds}sec"


def fetch_google_sheet_data():
    """Fetch data from the Google Sheet."""
    logging.info("Fetching data from Google Sheet...")
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SHEET_ID, range=RANGE_NAME).execute()
    data = result.get('values', [])
    logging.info("Google Sheet data fetched successfully.")
    return data, service, sheet

def update_google_sheet(service, sheet, cell_range, value):
    """Update the Google Sheet with processed data."""
    logging.info(f"Updating Google Sheet at {cell_range} with value: {value}")
    body = {'values': [[value]]}
    try:
        sheet.values().update(spreadsheetId=SHEET_ID, range=cell_range, valueInputOption="RAW", body=body).execute()
        logging.info("Google Sheet updated successfully.")
    except Exception as e:
        logging.error(f"Error updating Google Sheet: {e}")

#process and update the google sheet
def process_and_update_sheet():
    """Process the Google Sheet and update it with data from Close CRM."""
    sheet_data, service, sheet = fetch_google_sheet_data()
    headers = sheet_data[0]  # Extract header row (e.g., dates)

    for row_index, row in enumerate(sheet_data[1:], start=2):  # Skip header row (start=2 for actual row in Sheets)
        if len(row) > 2 and row[2]:  # Identify rows with email addresses
            email = row[2]
            logging.info(f"Processing data for email: {email}")

            # Locate the correct rows for "Call" and "Duration"
            call_row = None
            duration_row = None
            for i in range(row_index, len(sheet_data)):
                if sheet_data[i][0].lower() == "call":  # Locate "Call" row
                    call_row = i + 1
                if sheet_data[i][0].lower() == "duration":  # Locate "Duration" row
                    duration_row = i + 1
                if call_row and duration_row:
                    break

            if not call_row or not duration_row:
                logging.warning(f"Could not locate 'Call' or 'Duration' rows for email: {email}")
                continue

            for col_index, date in enumerate(headers[3:], start=3):  # Dates start from column D
                if date:
                    formatted_date = datetime.strptime(date, '%d/%m/%Y').strftime('%Y-%m-%d')
                    total_calls, total_duration = fetch_close_data(email, formatted_date, formatted_date)

                    # Update the "Call" row if necessary
                    call_cell_value = sheet_data[call_row - 1][col_index] if len(sheet_data[call_row - 1]) > col_index else ""
                    if not call_cell_value or not call_cell_value.isdigit() or int(call_cell_value) != total_calls:
                        call_cell_range = f"{chr(65 + col_index)}{call_row}"
                        update_google_sheet(service, sheet, call_cell_range, total_calls)

                    # Update the "Duration" row if necessary
                    formatted_duration = format_duration(total_duration)
                    duration_cell_value = sheet_data[duration_row - 1][col_index] if len(sheet_data[duration_row - 1]) > col_index else ""
                    if not duration_cell_value or duration_cell_value != formatted_duration:
                        duration_cell_range = f"{chr(65 + col_index)}{duration_row}"
                        update_google_sheet(service, sheet, duration_cell_range, formatted_duration)





if __name__ == "__main__":
    process_and_update_sheet()
