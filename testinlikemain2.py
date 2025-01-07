# update all google sheet cell but  data need to be checked backup 19/12-2024
from closeio_api import Client
from googleapiclient.discovery import build
from google.oauth2 import service_account
import logging
from datetime import datetime

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

 
SHEET_ID = "1-1rE5epVoYSJgqgkNYsGXuMrYQWkjbPjYpQUYgzMOTI"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'F:\\Giacometti\\New-Working-Directory-25-10-2024\\Python-report\\SERVICE_ACCOUNT_FILE\\closecrmdata-7f2f781a56e5.json'

CLOSE_API_KEY = "api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i"

 
api = Client(CLOSE_API_KEY)

 
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

 
def get_user_ids_by_email(email):
    all_users = fetch_all_team_members()
    return [user['id'] for user in all_users if user['email'] == email]

 
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

 
def fetch_sheet_names_from_settings():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SHEET_ID, range="Settings!B2:B").execute()
    sheet_names = [row[0] for row in result.get('values', []) if row]
    logging.info(f"All sheet names from the settings: {sheet_names}")
    return sheet_names

 
def fetch_google_sheet_data(sheet_name):
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SHEET_ID, range=f"{sheet_name}!A1:ZZZ").execute()
    data = result.get('values', [])
    return data, service, sheet

 
# previous working one --------------------------------------------------------------------
# def update_google_sheet(service, sheet, cell_range, value):
#     body = {'values': [[value]]}
#     try:
#         sheet.values().update(spreadsheetId=SHEET_ID, range=cell_range, valueInputOption="RAW", body=body).execute()
#     except Exception as e:
#         logging.error(f"Error updating Google Sheet at {cell_range}: {e}")


# if the cell is empty only then populate it otherwise do not populate it  . updated one 
# def update_google_sheet(service, sheet, cell_range, value):
    
#     try:
#         # Fetch the existing value of the cell
#         result = sheet.values().get(spreadsheetId=SHEET_ID, range=cell_range).execute()
#         existing_values = result.get("values", [[]])
#         existing_value = existing_values[0][0] if existing_values and existing_values[0] else ""

#         # Check if the cell is empty 
#         if not existing_value:  # Only update if the cell is empty
#             body = {'values': [[value]]}
#             sheet.values().update(
#                 spreadsheetId=SHEET_ID,
#                 range=cell_range,
#                 valueInputOption="RAW",
#                 body=body
#             ).execute()
#             logging.info(f"Updated {cell_range} with value: {value}")
#         else:
#             logging.info(f"Skipped updating {cell_range} as it already contains value: '{existing_value}'")
#     except Exception as e:
#         logging.error(f"Error updating Google Sheet at {cell_range}: {e}")


# if a cell value is empty or different then update it . ---------------------------------------------
def update_google_sheet(service, sheet, cell_range, value):
    try:
        # Fetch the existing value of the cell
        result = sheet.values().get(spreadsheetId=SHEET_ID, range=cell_range).execute()
        existing_values = result.get("values", [[]])
        existing_value = existing_values[0][0] if existing_values and existing_values[0] else ""

        # Decide whether to update
        if not existing_value:
            # Cell is empty, update with new value
            body = {'values': [[value]]}
            sheet.values().update(
                spreadsheetId=SHEET_ID,
                range=cell_range,
                valueInputOption="RAW",
                body=body
            ).execute()
            logging.info(f"Updated {cell_range} with value: {value} (was empty)")
        else:
            # Cell is not empty, check if the value is different
            if existing_value != str(value):
                # Existing value is different, update
                body = {'values': [[value]]}
                sheet.values().update(
                    spreadsheetId=SHEET_ID,
                    range=cell_range,
                    valueInputOption="RAW",
                    body=body
                ).execute()
                logging.info(f"Updated {cell_range} with new value: {value} (old value: {existing_value})")
            else:
                # Same value, skip updating
                logging.info(f"Skipped updating {cell_range} because it already contains the same value: '{existing_value}'")
    except Exception as e:
        logging.error(f"Error updating Google Sheet at {cell_range}: {e}")


 
def get_column_letter(index):
    column = ""
    while index >= 0:
        column = chr(index % 26 + 65) + column
        index = index // 26 - 1
    return column


# create a new method to update Discover prenotata cell in google sheet change starts
# Fetch Discovery Prenotata Data here

# Data e Ora Appuntamento  custom-field-id = cf_K5cgpMaCGXCPXd1c59Kosi8RvLBzz6DOQ09CdJSZgYk
def fetch_discovery_prenotata_data(assignee_name, date_filter, discovery_filter):
    print("Assignee name, date_filter, discovery_filter is -------------------:", assignee_name, date_filter, discovery_filter)
    limit = 100
    skip = 0
    all_leads = []
    matching_lead_names = []
    
    while True:
        try:
            response = api.get('lead', params={
                '_limit': limit,
                '_skip': skip,
                '_fields': 'id,display_name,opportunities,date_created',
                'query': f'custom.lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU:"{assignee_name}"'
            })
            all_leads.extend(response.get('data', []))
            if not response.get('has_more', False):
                break
            skip += limit
        except Exception as e:
            logging.error(f"Error fetching Discovery Prenotata data: {e}")
            break

    matching_leads = 0
    for lead in all_leads:
        for opportunity in lead.get('opportunities', []):
            lead_name = opportunity.get('lead_name')
            if (
                opportunity.get('status_display_name') == discovery_filter and 
                opportunity.get('date_created', '').startswith(date_filter) and 
                lead_name not in matching_lead_names
            ):
                matching_leads += 1
                matching_lead_names.append(lead_name)
                break

    print("Matching lead count -----------------------:", matching_leads)
    print("Matching lead names -----------------------:", matching_lead_names)
    return matching_leads

 

def fetch_discovery_programmata_data(assignee_name, date_filter, discovery_filter):
    print("Assignee name, date_filter, discovery_filter is -------------------:", assignee_name, date_filter, discovery_filter)
    limit = 100
    skip = 0
    all_leads = []
    matching_lead_names = []
    
    while True:
        try:
            response = api.get('lead', params={
                '_limit': limit,
                '_skip': skip,
                '_fields': 'id,display_name,opportunities,date_created',
                'query': f'custom.lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU:"{assignee_name}"'
            })
            all_leads.extend(response.get('data', []))
            if not response.get('has_more', False):
                break
            skip += limit
        except Exception as e:
            logging.error(f"Error fetching Discovery Prenotata data: {e}")
            break

    matching_leads = 0
    for lead in all_leads:
        for opportunity in lead.get('opportunities', []):
            lead_name = opportunity.get('lead_name')
            if (
                opportunity.get('status_display_name') == discovery_filter and 
                opportunity.get('date_created', '').startswith(date_filter) and 
                lead_name not in matching_lead_names
            ):
                matching_leads += 1
                matching_lead_names.append(lead_name)
                break

    print("Matching lead count -----------------------:", matching_leads)
    print("Matching lead names -----------------------:", matching_lead_names)
    return matching_leads


def fetch_discovery_completata_data(assignee_name, date_filter, discovery_filter):
    print("Assignee name, date_filter, discovery_filter is -------------------:", assignee_name, date_filter, discovery_filter)
    limit = 100
    skip = 0
    all_leads = []
    matching_lead_names = []
    
    while True:
        try:
            response = api.get('lead', params={
                '_limit': limit,
                '_skip': skip,
                '_fields': 'id,display_name,opportunities,date_created',
                'query': f'custom.lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU:"{assignee_name}"'
            })
            all_leads.extend(response.get('data', []))
            if not response.get('has_more', False):
                break
            skip += limit
        except Exception as e:
            logging.error(f"Error fetching Discovery Prenotata data: {e}")
            break

    matching_leads = 0
    for lead in all_leads:
        for opportunity in lead.get('opportunities', []):
            lead_name = opportunity.get('lead_name')
            if (
                opportunity.get('status_display_name') == discovery_filter and 
                opportunity.get('date_created', '').startswith(date_filter) and 
                #  opportunity.get('custom.cf_K5cgpMaCGXCPXd1c59Kosi8RvLBzz6DOQ09CdJSZgYk', '').startswith(date_filter) and
                lead_name not in matching_lead_names
            ):
                matching_leads += 1
                matching_lead_names.append(lead_name)
                break

    print("Matching lead count -----------------------:", matching_leads)
    print("Matching lead names -----------------------:", matching_lead_names)
    return matching_leads


def fetch_discovery_rischedulata_data(assignee_name, date_filter, discovery_filter):
    print("Assignee name, date_filter, discovery_filter is -------------------:", assignee_name, date_filter, discovery_filter)
    limit = 100
    skip = 0
    all_leads = []
    matching_lead_names = []
    
    while True:
        try:
            response = api.get('lead', params={
                '_limit': limit,
                '_skip': skip,
                '_fields': 'id,display_name,opportunities,date_created',
                'query': f'custom.lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU:"{assignee_name}"'
            })
            all_leads.extend(response.get('data', []))
            if not response.get('has_more', False):
                break
            skip += limit
        except Exception as e:
            logging.error(f"Error fetching Discovery Prenotata data: {e}")
            break

    matching_leads = 0
    for lead in all_leads:
        for opportunity in lead.get('opportunities', []):
            lead_name = opportunity.get('lead_name')
            if (
                opportunity.get('status_display_name') == discovery_filter and 
                opportunity.get('date_created', '').startswith(date_filter) and 
                lead_name not in matching_lead_names
            ):
                matching_leads += 1
                matching_lead_names.append(lead_name)
                break

    print("Matching lead count -----------------------:", matching_leads)
    print("Matching lead names -----------------------:", matching_lead_names)
    return matching_leads


def fetch_demo_prenotata_data(assignee_name, date_filter, discovery_filter):
    print("Assignee name, date_filter, discovery_filter is -------------------:", assignee_name, date_filter, discovery_filter)
    limit = 100
    skip = 0
    all_leads = []
    matching_lead_names = []
    
    while True:
        try:
            response = api.get('lead', params={
                '_limit': limit,
                '_skip': skip,
                '_fields': 'id,display_name,opportunities,date_created',
                'query': f'custom.lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU:"{assignee_name}"'
            })
            all_leads.extend(response.get('data', []))
            if not response.get('has_more', False):
                break
            skip += limit
        except Exception as e:
            logging.error(f"Error fetching Discovery Prenotata data: {e}")
            break

    matching_leads = 0
    for lead in all_leads:
        for opportunity in lead.get('opportunities', []):
            lead_name = opportunity.get('lead_name')
            if (
                opportunity.get('status_display_name') == discovery_filter and 
                opportunity.get('date_created', '').startswith(date_filter) and 
                lead_name not in matching_lead_names
            ):
                matching_leads += 1
                matching_lead_names.append(lead_name)
                break

    print("Matching lead count -----------------------:", matching_leads)
    print("Matching lead names -----------------------:", matching_lead_names)
    return matching_leads


def fetch_demo_programmata_data(assignee_name, date_filter, discovery_filter):
    print("Assignee name, date_filter, discovery_filter is -------------------:", assignee_name, date_filter, discovery_filter)
    limit = 100
    skip = 0
    all_leads = []
    matching_lead_names = []
    
    while True:
        try:
            response = api.get('lead', params={
                '_limit': limit,
                '_skip': skip,
                '_fields': 'id,display_name,opportunities,date_created',
                'query': f'custom.lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU:"{assignee_name}"'
            })
            all_leads.extend(response.get('data', []))
            if not response.get('has_more', False):
                break
            skip += limit
        except Exception as e:
            logging.error(f"Error fetching Discovery Prenotata data: {e}")
            break

    matching_leads = 0
    for lead in all_leads:
        for opportunity in lead.get('opportunities', []):
            lead_name = opportunity.get('lead_name')
            if (
                opportunity.get('status_display_name') == discovery_filter and 
                opportunity.get('date_created', '').startswith(date_filter) and 
                lead_name not in matching_lead_names
            ):
                matching_leads += 1
                matching_lead_names.append(lead_name)
                break

    print("Matching lead count -----------------------:", matching_leads)
    print("Matching lead names -----------------------:", matching_lead_names)
    return matching_leads


def fetch_demo_completata_data(assignee_name, date_filter, discovery_filter):
    print("Assignee name, date_filter, discovery_filter is -------------------:", assignee_name, date_filter, discovery_filter)
    limit = 100
    skip = 0
    all_leads = []
    matching_lead_names = []
    
    while True:
        try:
            response = api.get('lead', params={
                '_limit': limit,
                '_skip': skip,
                '_fields': 'id,display_name,opportunities,date_created',
                'query': f'custom.lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU:"{assignee_name}"'
            })
            all_leads.extend(response.get('data', []))
            if not response.get('has_more', False):
                break
            skip += limit
        except Exception as e:
            logging.error(f"Error fetching Discovery Prenotata data: {e}")
            break

    matching_leads = 0
    for lead in all_leads:
        for opportunity in lead.get('opportunities', []):
            lead_name = opportunity.get('lead_name')
            if (
                opportunity.get('status_display_name') == discovery_filter and 
                opportunity.get('date_created', '').startswith(date_filter) and 
                lead_name not in matching_lead_names
            ):
                matching_leads += 1
                matching_lead_names.append(lead_name)
                break

    print("Matching lead count -----------------------:", matching_leads)
    print("Matching lead names -----------------------:", matching_lead_names)
    return matching_leads


def fetch_demo_rischedulata_data(assignee_name, date_filter, discovery_filter):
    print("Assignee name, date_filter, discovery_filter is -------------------:", assignee_name, date_filter, discovery_filter)
    limit = 100
    skip = 0
    all_leads = []
    matching_lead_names = []
    
    while True:
        try:
            response = api.get('lead', params={
                '_limit': limit,
                '_skip': skip,
                '_fields': 'id,display_name,opportunities,date_created',
                'query': f'custom.lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU:"{assignee_name}"'
            })
            all_leads.extend(response.get('data', []))
            if not response.get('has_more', False):
                break
            skip += limit
        except Exception as e:
            logging.error(f"Error fetching Discovery Prenotata data: {e}")
            break

    matching_leads = 0
    for lead in all_leads:
        for opportunity in lead.get('opportunities', []):
            lead_name = opportunity.get('lead_name')
            if (
                opportunity.get('status_display_name') == discovery_filter and 
                opportunity.get('date_created', '').startswith(date_filter) and 
                lead_name not in matching_lead_names
            ):
                matching_leads += 1
                matching_lead_names.append(lead_name)
                break

    print("Matching lead count -----------------------:", matching_leads)
    print("Matching lead names -----------------------:", matching_lead_names)
    return matching_leads



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
            nome = row[0].strip() if len(row) > 0 else ""
            cognome = row[1].strip() if len(row) > 1 else ""
            assignee_name = f"{nome} {cognome}".strip()  # Combine NOME and COGNOME

            logging.info(f"Processing data for user: {assignee_name} (Email: {email}) in sheet: {sheet_name}")

            # Locate specific rows for this user
            outbound_calls_row, calls_duration_row, opportunity_won_row, value_won_annualized_row = None, None, None, None
            discovery_prenotata_row, discovery_programmata_row, discovery_completata_row, discovery_rischedulata_row = None, None, None, None
            demo_prenotata_row, demo_programmata_row, demo_completata_row, demo_rischedulata_row, leads_chiamati_row  = None, None, None, None,None



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
                    elif row_name == "discovery prenotata":
                        discovery_prenotata_row = i + 1
                    elif row_name == "discovery programmata":
                        discovery_programmata_row = i + 1
                    elif row_name == "discovery completata":
                        discovery_completata_row = i + 1
                    elif row_name == "discovery rischedulata":
                        discovery_rischedulata_row = i + 1
                    elif row_name == "demo prenotata":
                        demo_prenotata_row = i + 1
                    elif row_name == "demo programmata":
                        demo_programmata_row = i + 1
                    elif row_name == "demo completata":
                        demo_completata_row = i + 1
                    elif row_name == "demo rischedulata":
                        demo_rischedulata_row = i + 1
                    elif row_name == "leads chiamati":  # New addition
                        leads_chiamati_row = i + 1

                if all([outbound_calls_row, calls_duration_row, opportunity_won_row, value_won_annualized_row]):
                    break

            if not (outbound_calls_row and calls_duration_row and opportunity_won_row and value_won_annualized_row):
                logging.warning(f"Could not locate all required rows for user: {assignee_name}")
                continue

            # Update data for each date in the header row
            for col_index, date in enumerate(headers[3:], start=3):  # Dates start from column D
                column_letter = get_column_letter(col_index)
                logging.info(f"Processing date {date} at column {column_letter}")
                if date:
                    try:
                        formatted_date = datetime.strptime(date, '%d/%m/%Y').strftime('%Y-%m-%d')
                        total_calls, total_duration, lead_called = fetch_close_data(email, formatted_date, formatted_date)
                        total_opportunities, total_value = fetch_opportunity_data(email, formatted_date, formatted_date)

                        # Update the "Outbound Calls" row
                        calls_cell_range = f"{sheet_name}!{column_letter}{outbound_calls_row}"
                        update_google_sheet(service, sheet, calls_cell_range, total_calls)

                        # Update the "Calls Total Duration" row
                        formatted_duration = format_duration(total_duration)
                        duration_cell_range = f"{sheet_name}!{column_letter}{calls_duration_row}"
                        update_google_sheet(service, sheet, duration_cell_range, formatted_duration)

                        leads_chiamati_cell_range = f"{sheet_name}!{column_letter}{leads_chiamati_row}"
                        update_google_sheet(service, sheet, leads_chiamati_cell_range, lead_called)

                        # Update the "Opportunity Won" row
                        opportunity_cell_range = f"{sheet_name}!{column_letter}{opportunity_won_row}"
                        update_google_sheet(service, sheet, opportunity_cell_range, total_opportunities)

                        # Update the "Value Won Annualized" row
                        value_cell_range = f"{sheet_name}!{column_letter}{value_won_annualized_row}"
                        update_google_sheet(service, sheet, value_cell_range, total_value)

                        # Additional rows
                        if discovery_prenotata_row:
                            count = fetch_discovery_prenotata_data(assignee_name, formatted_date, "Discovery Programmata")
                            update_google_sheet(service, sheet, f"{sheet_name}!{column_letter}{discovery_prenotata_row}", count)

                        if discovery_programmata_row:
                            count = fetch_discovery_programmata_data(assignee_name, formatted_date, "Discovery Programmata")
                            update_google_sheet(service, sheet, f"{sheet_name}!{column_letter}{discovery_programmata_row}", count)

                        if discovery_completata_row:
                            count = fetch_discovery_completata_data(assignee_name, formatted_date, "Discovery Completata")
                            update_google_sheet(service, sheet, f"{sheet_name}!{column_letter}{discovery_completata_row}", count)

                        if discovery_rischedulata_row:
                            count = fetch_discovery_rischedulata_data(assignee_name, formatted_date, "Discovery Rischedulata")
                            update_google_sheet(service, sheet, f"{sheet_name}!{column_letter}{discovery_rischedulata_row}", count)

                        if demo_prenotata_row:
                            count = fetch_demo_prenotata_data(assignee_name, formatted_date, "Demo Prenotata")
                            update_google_sheet(service, sheet, f"{sheet_name}!{column_letter}{demo_prenotata_row}", count)

                        if demo_programmata_row:
                            count = fetch_demo_programmata_data(assignee_name, formatted_date, "Demo Programmata")
                            update_google_sheet(service, sheet, f"{sheet_name}!{column_letter}{demo_programmata_row}", count)

                        if demo_completata_row:
                            count = fetch_demo_completata_data(assignee_name, formatted_date, "Demo Completata")
                            update_google_sheet(service, sheet, f"{sheet_name}!{column_letter}{demo_completata_row}", count)

                        if demo_rischedulata_row:
                            count = fetch_demo_rischedulata_data(assignee_name, formatted_date, "Demo Rischedulata")
                            update_google_sheet(service, sheet, f"{sheet_name}!{column_letter}{demo_rischedulata_row}", count)

                    except Exception as e:
                        logging.error(f"Error processing date {date} for user {assignee_name} in sheet {sheet_name}: {e}")
                        continue



def main():
    sheet_names = fetch_sheet_names_from_settings()
    for sheet_name in sheet_names:
        logging.info(f"Starting processing for sheet: {sheet_name}")
        process_and_update_sheet(sheet_name)
    logging.info("Processing completed for all sheets.")

# if __name__ == "__main__":
#     main()

import time

if __name__ == "__main__":
    while True:
        main()
        print ("-------------------------------check again the google sheet in 5 sec-------------------------------------------------")
        time.sleep(5)  # Wait for 5 seconds before running main() again


















# previous working one  
# def process_and_update_sheet(sheet_name):
#     logging.info(f"Processing sheet: {sheet_name}")
#     sheet_data, service, sheet = fetch_google_sheet_data(sheet_name)
#     if not sheet_data:
#         logging.warning(f"Sheet {sheet_name} is empty or unavailable.")
#         return

#     headers = sheet_data[0]   
#     logging.info(f"Header row for sheet {sheet_name}: {headers}")

#     for row_index, row in enumerate(sheet_data[1:], start=2):  
#         if len(row) > 2 and row[2]:  
#             email = row[2]
#             logging.info(f"Processing data for email: {email} in sheet: {sheet_name}")

         
#             outbound_calls_row, calls_duration_row, opportunity_won_row, value_won_annualized_row = None, None, None, None
#             for i in range(row_index, len(sheet_data)):
#                 if len(sheet_data[i]) > 0:
#                     row_name = sheet_data[i][0].strip().lower()
#                     if row_name == "outbound calls":
#                         outbound_calls_row = i + 1
#                     elif row_name == "calls total duration":
#                         calls_duration_row = i + 1
#                     elif row_name == "opportunity won":
#                         opportunity_won_row = i + 1
#                     elif row_name == "value won annualized":
#                         value_won_annualized_row = i + 1
#                 if outbound_calls_row and calls_duration_row and opportunity_won_row and value_won_annualized_row:
#                     break

#             if not (outbound_calls_row and calls_duration_row and opportunity_won_row and value_won_annualized_row):
#                 logging.warning(f"Could not locate required rows for email: {email}")
#                 continue

       
#             for col_index, date in enumerate(headers[3:], start=3):   
#                 column_letter = get_column_letter(col_index)
#                 logging.info(f"Processing date {date} at column {column_letter}")
#                 if date:
#                     try:
#                         formatted_date = datetime.strptime(date, '%d/%m/%Y').strftime('%Y-%m-%d')
#                         total_calls, total_duration, _ = fetch_close_data(email, formatted_date, formatted_date)
#                         total_opportunities, total_value = fetch_opportunity_data(email, formatted_date, formatted_date)

                    

#                         # create a new mecanism to call fetch_discovery_prenotata_data here 
#                         #  here pass assignee_name and  date after decode google sheet  like other method                

#                         calls_cell_range = f"{sheet_name}!{column_letter}{outbound_calls_row}"
#                         update_google_sheet(service, sheet, calls_cell_range, total_calls)

          
#                         formatted_duration = format_duration(total_duration)
#                         duration_cell_range = f"{sheet_name}!{column_letter}{calls_duration_row}"
#                         update_google_sheet(service, sheet, duration_cell_range, formatted_duration)

                
#                         opportunity_cell_range = f"{sheet_name}!{column_letter}{opportunity_won_row}"
#                         update_google_sheet(service, sheet, opportunity_cell_range, total_opportunities)

                  
#                         value_cell_range = f"{sheet_name}!{column_letter}{value_won_annualized_row}"
#                         update_google_sheet(service, sheet, value_cell_range, total_value)

#                         logging.info(f"Updated {calls_cell_range} with {total_calls}")
#                         logging.info(f"Updated {duration_cell_range} with {formatted_duration}")
#                         logging.info(f"Updated {opportunity_cell_range} with {total_opportunities}")
#                         logging.info(f"Updated {value_cell_range} with {total_value}")

#                     except Exception as e:
#                         logging.error(f"Error processing date {date} for email {email} in sheet: {sheet_name}: {e}")
#                         continue
