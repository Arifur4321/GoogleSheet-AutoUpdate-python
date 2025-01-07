import json
from closeio_api import Client

# Initialize the Close.io API client
api = Client('api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i')

# Variables to test
assignee_name = "Simone Banfi"
custom_field_id = "lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU"

# Fetch metadata for custom fields to ensure the ID is valid
print("Fetching custom field metadata...")
custom_fields_response = api.get('custom_fields/lead')

if 'data' in custom_fields_response:
    custom_fields = custom_fields_response['data']
    field_metadata = next(
        (field for field in custom_fields if field['id'] == custom_field_id),
        None
    )

    if field_metadata:
        print(f"Custom Field Found: {field_metadata['name']} (ID: {field_metadata['id']})")
    else:
        print(f"Custom Field ID {custom_field_id} not found.")
        exit()
else:
    print("Error fetching custom fields.")
    exit()

# Pagination variables
skip = 0
limit = 100  # Number of leads to fetch per request
has_more = True

# Debugging variables
total_leads_checked = 0
matching_leads = 0

while has_more:
    # Fetch leads from Close.io
    lead_results = api.get('lead', params={
        '_limit': limit,
        '_skip': skip,
        '_fields': f'id,display_name,custom.{custom_field_id}'
    })

    for lead in lead_results['data']:
        total_leads_checked += 1
        
        # Ensure 'custom' key exists before accessing
        assignee_field_value = lead.get('custom', {}).get(custom_field_id, None)

        # Print debug information
        print(f"Lead ID: {lead['id']}, Name: {lead['display_name']}, Assignee Field Value: {assignee_field_value}")

        # Check if the custom field value matches the given assignee name
        if assignee_field_value == assignee_name:
            matching_leads += 1
            print(f"Match found for Lead ID: {lead['id']} (Name: {lead['display_name']})")

    # Update pagination variables
    has_more = lead_results.get('has_more', False)
    skip += limit

# Print summary
print("\nSummary")
print(f"Total Leads Checked: {total_leads_checked}")
print(f"Total Matching Leads: {matching_leads}")
