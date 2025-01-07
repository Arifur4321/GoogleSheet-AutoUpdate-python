# -------------------------------------------list all smart view with id -----------------------------------------------
# from closeio_api import Client

# # Replace with your Close API key
# CLOSE_API_KEY = "api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i"
# api = Client(CLOSE_API_KEY)

# # Function to list all Smart Views
# def list_smart_views():
#     try:
#         response = api.get('saved_search/')
#         smart_views = response.get('data', [])
#         for smart_view in smart_views:
#             print(f"Name: {smart_view['name']}, ID: {smart_view['id']}")
#     except Exception as e:
#         print(f"Error fetching Smart Views: {e}")

# # Call the function
# list_smart_views()
#------------------------------------------------------------------------------------------------------

# from closeio_api import Client

# # Replace with your Close API key
# CLOSE_API_KEY = "api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i"
# api = Client(CLOSE_API_KEY)

# # Specify the Smart View ID
# smart_view_id = 'save_vYgs1eluT1Fz2mLMqCK5aYZgnFexFwOpuztDHoj7bWH'

# # Specify the params to fetch only the query
# params = {"_fields": "query"}

# # Use an f-string to insert the Smart View ID into the endpoint URL
# resp = api.get(f'saved_search/{smart_view_id}', params=params)

# # Print the response user_id = 'user_HRPX286qNrByGXLMxzRcRwlWBui3mOEe7dZ4bs0vV1k'
# print(resp)
#------------------------------------------------------------------------------------------------------------------------

# from closeio_api import Client

# # Constants
# CLOSE_API_KEY = "api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i"
# SMART_VIEW_ID = "save_vYgs1eluT1Fz2mLMqCK5aYZgnFexFwOpuztDHoj7bWH"
# USER_ID = "user_HRPX286qNrByGXLMxzRcRwlWBui3mOEe7dZ4bs0vV1k"
# TARGET_DATE = "2024-12-02"  # Using the format from JSON: YYYY-MM-DD

# # Initialize the Close API client
# api = Client(CLOSE_API_KEY)

# def fetch_leads_with_filters(api_client, smart_view_id, date):
#     """
#     Fetches leads based on the filters extracted from the JSON.

#     Args:
#         api_client (Client): Instance of the Close API client.
#         smart_view_id (str): The ID of the smart view.
#         date (str): The date of interest (format: YYYY-MM-DD).

#     Returns:
#         list: A list of dictionaries containing lead information.
#     """
#     params = {
#         "query": {
#             "negate": False,
#             "queries": [
#                 {
#                     "negate": False,
#                     "object_type": "lead",
#                     "type": "object_type"
#                 },
#                 {
#                     "negate": False,
#                     "queries": [
#                         {
#                             "negate": False,
#                             "related_object_type": "opportunity",
#                             "related_query": {
#                                 "negate": False,
#                                 "queries": [
#                                     {
#                                         "condition": {
#                                             "before": {
#                                                 "type": "fixed_local_date",
#                                                 "value": date,
#                                                 "which": "end"
#                                             },
#                                             "on_or_after": {
#                                                 "type": "fixed_local_date",
#                                                 "value": date,
#                                                 "which": "start"
#                                             },
#                                             "type": "moment_range"
#                                         },
#                                         "field": {
#                                             "field_name": "date_created",
#                                             "object_type": "opportunity",
#                                             "type": "regular_field"
#                                         },
#                                         "negate": False,
#                                         "type": "field_condition"
#                                     },
#                                     {
#                                         "before": {
#                                             "type": "fixed_local_date",
#                                             "value": date,
#                                             "which": "end"
#                                         },
#                                         "negate": False,
#                                         "on_or_after": {
#                                             "type": "fixed_local_date",
#                                             "value": date,
#                                             "which": "start"
#                                         },
#                                         "status_ids": [
#                                             "stat_lMssygbZ8KyDSxe7QfcAYOxhsfb0r0RmxBsPgD5oSDB"
#                                         ],
#                                         "type": "in_status"
#                                     }
#                                 ],
#                                 "type": "and"
#                             },
#                             "this_object_type": "lead",
#                             "type": "has_related"
#                         },
#                         {
#                             "negate": False,
#                             "queries": [
#                                 {
#                                     "condition": {
#                                         "mode": "beginning_of_words",
#                                         "type": "text",
#                                         "value": "Simone Banfi"
#                                     },
#                                     "field": {
#                                         "custom_field_id": "lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU",
#                                         "type": "custom_field"
#                                     },
#                                     "negate": False,
#                                     "type": "field_condition"
#                                 }
#                             ],
#                             "type": "and"
#                         }
#                     ],
#                     "type": "and"
#                 }
#             ],
#             "type": "and"
#         }
#     }
#     resp = api_client.get(f'saved_search/{smart_view_id}', params=params)
#     print("all response data-----------------:", resp)
#     if 'data' in resp:
#         return resp['data']
#     else:
#         return []

# def print_lead_info(leads):
#     """
#     Prints detailed information for each lead.

#     Args:
#         leads (list): List of lead dictionaries.
#     """
#     if not leads:
#         print("No leads found for the specified filters.")
#         return

#     print(f"Leads based on filters for {TARGET_DATE}:")
#     for lead in leads:
#         display_name = lead.get("display_name", "N/A")
#         primary_phone = lead.get("primary_phone", "N/A")
#         primary_email = lead.get("primary_email", "N/A")
#         primary_contact_name = lead.get("primary_contact_name", "N/A")
#         status_id = lead.get("status_id", "N/A")
#         custom_field = lead.get("lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU", "N/A")

#         print(f"Lead Name: {display_name}")
#         print(f"  Phone: {primary_phone}")
#         print(f"  Email: {primary_email}")
#         print(f"  Primary Contact: {primary_contact_name}")
#         print(f"  Status ID: {status_id}")
#         print(f"  Custom Field: {custom_field}")
#         print("-" * 40)

# # Main execution
# if __name__ == "__main__":
#     try:
#         # Fetch leads based on filters
#         leads = fetch_leads_with_filters(api, SMART_VIEW_ID, TARGET_DATE)
#         # Print lead details
#         print_lead_info(leads)
#     except Exception as e:
#         print(f"Error: {e}")

#------------------------------------------------------------------------------------------

# from closeio_api import Client

# # Constants
# CLOSE_API_KEY = "api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i"
# USER_ID = "user_HRPX286qNrByGXLMxzRcRwlWBui3mOEe7dZ4bs0vV1k"
# TARGET_DATE = "2024-12-03"  # Format: YYYY-MM-DD

# # Initialize the Close API client
# api = Client(CLOSE_API_KEY)

# def fetch_leads(api_client, user_id, date):
#     """
#     Fetches leads based on the user ID and date.
#     """
#     # Start with a basic query
#     params = {
#         "query": f"lead.user_id:{user_id} AND lead.date_created:{date}"
#     }
#     print(f"Query: {params}")
#     try:
#         response = api_client.get("lead", params=params)
#         print("Response:", response)
#         return response.get("data", [])
#     except Exception as e:
#         print(f"Error fetching leads: {e}")
#         return []

# def print_leads(leads):
#     """
#     Prints details of the fetched leads.
#     """
#     if not leads:
#         print("No leads found for the given filters.")
#         return

#     print(f"Leads for user on {TARGET_DATE}:")
#     for lead in leads:
#         display_name = lead.get("display_name", "N/A")
#         primary_phone = lead.get("primary_phone", "N/A")
#         primary_email = lead.get("primary_email", "N/A")
#         primary_contact_name = lead.get("primary_contact_name", "N/A")
#         status_id = lead.get("status_id", "N/A")

#         print(f"Lead Name: {display_name}")
#         print(f"  Phone: {primary_phone}")
#         print(f"  Email: {primary_email}")
#         print(f"  Primary Contact: {primary_contact_name}")
#         print(f"  Status ID: {status_id}")
#         print("-" * 40)

# # Main execution
# if __name__ == "__main__":
#     try:
#         # Fetch leads
#         leads = fetch_leads(api, USER_ID, TARGET_DATE)
#         # Print lead details
#         print_leads(leads)
#     except Exception as e:
#         print(f"Error: {e}")
#-----------------------------------------------------------------------

# params = {
#     '_fields': 'id,name,contacts,opportunities,date_created,date_updated',  # Fields to include in the response
#     'date_created__gte': f'{date}T00:00:00Z',  # Start of December 3, 2024
#     'date_created__lte': f'{date}T23:59:59Z'   # End of December 3, 2024
# }


# so can you modify the python code to get the result as in the screenshot
#-------------------------------------------------------------------------

# import json
# from closeio_api import Client

# # Initialize the Close.io API client
# api = Client('api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i')

# # Pagination variables
# skip = 0
# limit = 100
# has_more = True

# # Define the assignee name and filter criteria
# assignee_name = "Simone Banfi"
# status_name = "Discovery Programmata"

# # Container for all filtered leads
# filtered_leads = []

# # Fetch and process leads with pagination
# while has_more:
#     # Adding more fields to `_fields` for detailed data
#     lead_results = api.get('lead', params={
#         '_limit': limit,
#         '_skip': skip,
#         '_fields': 'id,display_name,opportunities,custom,contacts,created_by_name,updated_by_name,date_created,date_updated',
#         'query': f'custom.lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU:"{assignee_name}"'
#     })

#     # Process each lead
#     for lead in lead_results['data']:
#         print(f"\nLead ID: {lead['id']}, Name: {lead['display_name']}")
#         if 'opportunities' in lead:
#             for opp in lead['opportunities']:
#                 # Check if `status_display_name` matches "Discovery Programmata"
#                 if opp.get('status_display_name') == status_name:
#                     filtered_leads.append({
#                         "lead_id": lead['id'],
#                         "lead_name": lead['display_name'],
#                         "opportunity_id": opp['id'],
#                         "opportunity_status": opp['status_display_name'],
#                         "date_created": opp.get('date_created'),
#                         "updated_by_name": lead.get('updated_by_name'),
#                         "created_by_name": lead.get('created_by_name')
#                     })
#                     print(json.dumps(opp, indent=4))  # Print the opportunity data

#     # Update pagination variables
#     has_more = lead_results.get('has_more', False)
#     skip += limit

# # Print filtered leads
# print("\nFiltered Leads:")
# print(json.dumps(filtered_leads, indent=4))
#----------------------------------------------       Almost Working      -----------------------------------------------------------------------

# import json
# from closeio_api import Client

# # Initialize the Close.io API client
# api = Client('api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i')

# # Pagination variables
# skip = 0
# limit = 100
# has_more = True

# # Define the criteria
# assignee_name = "Simone Banfi"
# booking_date = "2024-12-03"  # Booking date to filter
# filtered_leads = []  # Container for filtered leads
# #opportunity_status = "Discovery Programmata" 

# # Fetch and process leads with pagination
# while has_more:
#     # Adding more fields to `_fields` for detailed data
#     lead_results = api.get('lead', params={
#         '_limit': limit,
#         '_skip': skip,
#         '_fields': 'id,display_name,opportunities,custom,contacts,created_by_name,updated_by_name,date_created,date_updated',
#         'query': f'custom.lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU:"{assignee_name}"'
#     })

#     # Process each lead
#     for lead in lead_results['data']:
#         if 'opportunities' in lead:
#             for opp in lead['opportunities']:
#                 # Filter opportunities by booking date and assignee
#                 date_created = opp.get('date_created', "")
#                 opp_status = opp.get('status_display_name', "")
#                 if booking_date in date_created and opp.get('user_name') == assignee_name:
#                     filtered_leads.append({
#                         "lead_id": lead['id'],
#                         "lead_name": lead['display_name'],
#                         "date_created": date_created,
#                         "opportunity_id": opp['id'],
#                         "assigned_to": assignee_name,
#                         "opportunity_status": opp_status
#                     })

#     # Update pagination variables
#     has_more = lead_results.get('has_more', False)
#     skip += limit

# # Print the filtered leads
# print("\nFiltered Leads:")
# print(json.dumps(filtered_leads, indent=4))

#---------------------------------------  fetch all leads with contact and opportunity asigned to = Simone Banfi  -----------------------------------------------------------------------------------

import json
from closeio_api import Client

# Initialize Close API client with your API key
api = Client('api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i')

# Define parameters
assignee_name = "Alessio Chianetta"  # Replace with the actual assignee name
limit = 100  # Adjust limit for the number of leads to fetch per request
skip = 0  # Pagination starts from 0
all_leads = []  # To store all the fetched leads

while True:
    # Fetch leads with the specified query and fields
    response = api.get('lead', params={
        '_limit': limit,
        '_skip': skip,
        '_fields': 'id,display_name,opportunities,custom,contacts,created_by_name,updated_by_name,date_created,date_updated',
        'query': f'custom.lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU:"{assignee_name}"'
    })
    
    # Add fetched leads to the list
    all_leads.extend(response.get('data', []))
    
    # Check if there are more leads to fetch
    if not response.get('has_more', False):
        break  # Exit loop if no more leads to fetch
    
    # Update skip for pagination
    skip += limit

# Write all leads to a JSON file
with open('filtered_leads.json', 'w', encoding='utf-8') as file:
    json.dump(all_leads, file, indent=4)

# Print the response in a human-readable format
print("Fetched Leads:")
for lead in all_leads:
    print(json.dumps(lead, indent=4))

 

#-------------------------------------------- for getting all custom field id ----------------------------------------------------- 
# import json
# from closeio_api import Client

# # Initialize Close API client with your API key
# api = Client('api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i')

# # Define any parameters you need for filtering custom fields
# # For example, if no filters are needed, you can leave it as an empty dict.
# params = {
#     # 'some_parameter': 'some_value'
# }

# # Fetch the custom fields
# resp = api.get('custom_field/shared', params=params)

# # Print the response in a human-readable way (JSON with indentation)
# print("Human-readable custom fields response:")
# print(json.dumps(resp, indent=4))

# # Write the entire response to a JSON file
# with open('filtered_leads.json', 'w', encoding='utf-8') as f:
#     json.dump(resp, f, indent=4)

# print("Custom fields information has been written to filtered_leads.json.")


#----------------------------------using filter to get leads-------------------------------------------------------------------------

# import json
# from closeio_api import Client

# # Initialize Close API client with your API key
# api = Client('api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i')

# # Define parameters
# assignee_name = "Simone Banfi"   
# date_filter = "2024-12-02"   
# discovery_filter = "Discovery Programmata"   
# limit = 100  # Adjust limit for the number of leads to fetch per request
# skip = 0  # Pagination starts from 0
# all_leads = []  # To store all the fetched leads

# while True:
#     # Fetch leads with the specified query and fields
#     response = api.get('lead', params={
#         '_limit': limit,
#         '_skip': skip,
#         '_fields': 'id,display_name,opportunities,custom,contacts,created_by_name,updated_by_name,date_created,date_updated',
#         'query': f'custom.lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU:"{assignee_name}"'
#     })

#     # Add fetched leads to the list
#     all_leads.extend(response.get('data', []))

#     # Check if there are more leads to fetch
#     if not response.get('has_more', False):
#         break  # Exit loop if no more leads to fetch

#     # Update skip for pagination
#     skip += limit

# # Filter opportunities with the specific status and date
# matching_leads = set()

# for lead in all_leads:
#     if 'opportunities' in lead:
#         for opportunity in lead['opportunities']:
#             if (opportunity.get('status_display_name') == discovery_filter and
#                 opportunity.get('status_label') == discovery_filter and
#                 opportunity.get('date_created', '').startswith(date_filter)):
#                 matching_leads.add(lead['display_name'])
#                 break  # Avoid duplicate leads when multiple opportunities match

# # Write all leads to a JSON file (optional for reference)
# with open('filtered_leads.json', 'w', encoding='utf-8') as file:
#     json.dump(all_leads, file, indent=4)

# # Print the filtered lead names and count
# print(f"Leads with '{discovery_filter}' opportunities on {date_filter}:")
# for lead_name in sorted(matching_leads):
#     print(lead_name)

# print(f"Total matching leads: {len(matching_leads)}")

 

#-------------------------------------------------------- get all user id and email of team members *----------------------------------------------------------------

# import json
# import logging
# from closeio_api import Client

# # Initialize Close API client
# api = Client('api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i')

# # Configure logging
# logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# def fetch_all_team_members():
 
#     all_users = []
#     has_more = True
#     offset = 0
#     limit = 100

#     while has_more:
#         try:
#             # Fetch team members with pagination
#             response = api.get('user/', params={'_skip': offset, '_limit': limit})
#             users = response.get('data', [])
#             all_users.extend(users)
            
#             # Check for more data
#             has_more = response.get('has_more', False)
#             offset += limit
#         except Exception as e:
#             logging.error(f"Error fetching team members: {e}")
#             has_more = False  # Stop fetching on error

#     return [
#         {
#             "id": user.get("id"),
#             "email": user.get("email", "N/A"),
#         }
#         for user in all_users
#     ]

# def save_to_json_file(data, filename):
   
#     try:
#         with open(filename, 'w', encoding='utf-8') as file:
#             json.dump(data, file, indent=4)
#         logging.info(f"Data successfully saved to {filename}")
#     except IOError as e:
#         logging.error(f"Failed to save data to {filename}: {e}")

# def print_and_save_user_ids():
    
#     logging.info("Fetching all team members...")
#     all_users = fetch_all_team_members()

#     if not all_users:
#         logging.warning("No team members found.")
#         return

#     # Print users to console
#     logging.info("Fetched team members successfully:")
#     for user in all_users:
#         print(f"User ID: {user['id']} | Email: {user['email']}")

#     # Save to JSON file
#     save_to_json_file(all_users, "filtered_leads.json")

# if __name__ == "__main__":
#     # Fetch, print, and save all team member IDs
#     print_and_save_user_ids()

#---------------------------------------- testing purpose -------------------------------------------------------------------------

# from closeio_api import Client
# import json

# # Initialize the Close API client with your API key
# api = Client('api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i')

# params = {}
# # Fetch custom object types
# try:
#     # Perform the GET request
#     response = api.get('custom_field/lead', params=params)

#     # Format the response as JSON for readability
#     formatted_response = json.dumps(response, indent=4)

#     # Print the response in human-readable format
#     print("Custom Object Types:")
#     print(formatted_response)

#     # Save the response to a JSON file
#     with open('filtered_leads.json', 'w') as file:
#         json.dump(response, file, indent=4)
#         print("Response has been saved to 'filtered_leads.json'.")

# except Exception as e:
#     print(f"An error occurred while fetching data: {e}")


#-----------------------------------------------------testing code-------------------------------------------------------------------
# from closeio_api import Client
# import json

# API_KEY = 'api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i'  # Replace with your Close API key
# TEST_USER_ID = 'user_HRPX286qNrByGXLMxzRcRwlWBui3mOEe7dZ4bs0vV1k'  # Replace with a valid Close user ID of simone banfi

# api = Client(API_KEY)

# params = {
#   #  "query": f"assigned_to.id:{TEST_USER_ID}"
#      "query": f"assigned_to.id:{TEST_USER_ID} AND status_label:Active"     
# }

# try:
#     response = api.get('opportunity', params=params)
#     print ("latest simone banfi opportunity : ", response)
#     opportunities = response.get('data', [])
#     print(f"Found {len(opportunities)} opportunities assigned to user {TEST_USER_ID}.")

#     # Print each opportunity's ID and lead name for verification
#     for opp in opportunities:
#         opp_id = opp.get('id')
#         lead_name = opp.get('lead_name', 'Unknown Lead')
#         print(f"Opportunity ID: {opp_id}, Lead Name: {lead_name}")

#     # Optionally, print the full JSON response:
#     # print(json.dumps(response, indent=4))

# except Exception as e:
#     print(f"An error occurred: {e}")
