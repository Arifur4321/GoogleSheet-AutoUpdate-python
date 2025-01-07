# works fine upto opportunity won and value won annulaized
from closeio_api import Client
from googleapiclient.discovery import build
from google.oauth2 import service_account
import logging
from datetime import datetime

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

# Function to fetch Close CRM data for activities
def fetch_close_data(email, start_date, end_date):
    user_ids = get_user_ids_by_email(email)
    if not user_ids:
        logging.warning(f"No user IDs found for email {email}. Skipping.")
        return 0, 0, 0

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
                "calls.outbound.all.sum_duration",
                "leads.contacted.all.count"
            ],
        }
        response = api.post('report/activity', data=data)
        metrics = response.get("aggregations", {}).get("totals", {})
        total_calls = metrics.get("calls.outbound.all.count", 0)
        total_duration = metrics.get("calls.outbound.all.sum_duration", 0)
        leads_created = metrics.get("leads.contacted.all.count", 0)
        return total_calls, total_duration, leads_created
    except Exception as e:
        logging.error(f"Error fetching Close CRM data for {email}: {e}")
        return 0, 0, 0

# Function to fetch Close CRM data for opportunities
def fetch_opportunity_data(email, start_date, end_date):
    user_ids = get_user_ids_by_email(email)
    if not user_ids:
        logging.warning(f"No user IDs found for email {email}. Skipping.")
        return 0, 0

    try:
        params = {
            "query": f"date_created >= {start_date} AND date_created <= {end_date} AND assigned_to.id:{','.join(user_ids)}",
        }
        response = api.get('opportunity/', params=params)
        opportunities = response.get('data', [])
        total_opportunities = len(opportunities)
        total_value = sum(opportunity.get("value", 0) for opportunity in opportunities)
        return total_opportunities, total_value
    except Exception as e:
        logging.error(f"Error fetching opportunities data for {email}: {e}")
        return 0, 0

# Function to format duration in seconds
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

# Function to fetch sheet names from the "Settings" sheet
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
    result = sheet.values().get(spreadsheetId=SHEET_ID, range=f"{sheet_name}!A1:ZZZ").execute()
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
    column = ""
    while index >= 0:
        column = chr(index % 26 + 65) + column
        index = index // 26 - 1
    return column

# Function to process and update a sheet
def process_and_update_sheet(sheet_name):
    logging.info(f"Processing sheet: {sheet_name}")
    sheet_data, service, sheet = fetch_google_sheet_data(sheet_name)
    if not sheet_data:
        logging.warning(f"Sheet {sheet_name} is empty or unavailable.")
        return

    headers = sheet_data[0]  # Extract header row (e.g., dates)
    logging.info(f"Header row for sheet {sheet_name}: {headers}")

    for row_index, row in enumerate(sheet_data[1:], start=2):  # Skip header row
        if len(row) > 2 and row[2]:  # Identify rows with email addresses
            email = row[2]
            logging.info(f"Processing data for email: {email} in sheet: {sheet_name}")

            # Locate specific rows for this user
            outbound_calls_row, calls_duration_row, opportunity_won_row, value_won_annualized_row = None, None, None, None
            for i in range(row_index, len(sheet_data)):
                if len(sheet_data[i]) > 0:
                    row_name = sheet_data[i][0].strip().lower()
                    if row_name == "outbound calls":
                        outbound_calls_row = i + 1
                    elif row_name == "calls total duration":
                        calls_duration_row = i + 1
                    elif row_name == "opportunity won":
                        opportunity_won_row = i + 1
                    elif row_name == "value won annualized":
                        value_won_annualized_row = i + 1
                if outbound_calls_row and calls_duration_row and opportunity_won_row and value_won_annualized_row:
                    break

            if not (outbound_calls_row and calls_duration_row and opportunity_won_row and value_won_annualized_row):
                logging.warning(f"Could not locate required rows for email: {email}")
                continue

            # Update data for each date in the header row
            for col_index, date in enumerate(headers[3:], start=3):  # Dates start from column D
                column_letter = get_column_letter(col_index)
                logging.info(f"Processing date {date} at column {column_letter}")
                if date:
                    try:
                        formatted_date = datetime.strptime(date, '%d/%m/%Y').strftime('%Y-%m-%d')
                        total_calls, total_duration, _ = fetch_close_data(email, formatted_date, formatted_date)
                        total_opportunities, total_value = fetch_opportunity_data(email, formatted_date, formatted_date)

                        # Update the "Outbound Calls" row

                        calls_cell_range = f"{sheet_name}!{column_letter}{outbound_calls_row}"
                        update_google_sheet(service, sheet, calls_cell_range, total_calls)

                        # Update the "Calls Total Duration" row
                        formatted_duration = format_duration(total_duration)
                        duration_cell_range = f"{sheet_name}!{column_letter}{calls_duration_row}"
                        update_google_sheet(service, sheet, duration_cell_range, formatted_duration)

                        # Update the "Opportunity Won" row
                        opportunity_cell_range = f"{sheet_name}!{column_letter}{opportunity_won_row}"
                        update_google_sheet(service, sheet, opportunity_cell_range, total_opportunities)

                        # Update the "Value Won Annualized" row
                        value_cell_range = f"{sheet_name}!{column_letter}{value_won_annualized_row}"
                        update_google_sheet(service, sheet, value_cell_range, total_value)

                        logging.info(f"Updated {calls_cell_range} with {total_calls}")
                        logging.info(f"Updated {duration_cell_range} with {formatted_duration}")
                        logging.info(f"Updated {opportunity_cell_range} with {total_opportunities}")
                        logging.info(f"Updated {value_cell_range} with {total_value}")

                    except Exception as e:
                        logging.error(f"Error processing date {date} for email {email} in sheet: {sheet_name}: {e}")
                        continue

# Main function to process all sheets
def main():
    sheet_names = fetch_sheet_names_from_settings()
    for sheet_name in sheet_names:
        logging.info(f"Starting processing for sheet: {sheet_name}")
        process_and_update_sheet(sheet_name)
    logging.info("Processing completed for all sheets.")

if __name__ == "__main__":
    main()
