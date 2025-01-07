import json
from closeio_api import Client

# Initialize the Close.io API client with your API key
api = Client('api_33gDN2iQffFOdrOEPtsUdg.3lzKAcaZDseZ4MblVHMN7i')

 

# Set your object_type
object_type = 'lead'  # Adjust if necessary

# Define the data for the PUT request
# Add all field IDs you want to check
field_ids = ["lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU"]

data = {
    "fields": [{"id": field_id} for field_id in field_ids]
}

# Make the PUT request to update the custom field schema
resp = api.put(f'custom_field_schema/{object_type}', data=data)

# Check the response for the target custom field
response_data = resp.get("fields", [])
target_custom_field_id = "lcf_NVYM9Lbe5754j18xHbmbbITZR1OArJqNVsosiyXhGKU"

# Search for the target custom field
found_field = next((field for field in response_data if field.get("id") == target_custom_field_id), None)

# Print the result in a human-readable format
if found_field:
    print("Custom Field Found:")
    print(json.dumps(found_field, indent=4, sort_keys=True))
else:
    print(f"Custom Field with ID '{target_custom_field_id}' not found.")

