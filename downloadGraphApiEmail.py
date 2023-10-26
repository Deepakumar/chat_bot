from promptflow import tool
import requests
import json

# The inputs section will change based on the arguments of the tool function, after you save the code
# Adding type to arguments and return value will help the system show the types properly
# Please update the function name/signature per need
@tool
def my_python_tool(input1: str) -> str:

  # Read the configuration file
  with open('config.json', 'r') as config_file:
    config = json.load(config_file)

  # Read Grapy API Configuration
  user_id = config['user_id']
  graph_api_endpoint = f"{config['graph_api_endpoint'].format(user_id=user_id)}"
  client_id = config['client_id']
  client_secret = config['client_secret']
  redirect_uri = config['redirect_uri']
  tenant_id = config['tenant_id']
  access_token = config['access_token']

  # Print the configuration
  print("Configuration:")
  print(f"User ID: {user_id}")
  print(f"Graph API Endpoint: {graph_api_endpoint}")
  print(f"Client ID: {client_id}")
  print(f"Client Secret: {client_secret}")
  print(f"Redirect URI: {redirect_uri}")
  print(f"Tenant ID: {tenant_id}")    

  get_daily_emails(graph_api_endpoint, client_id, client_secret, redirect_uri, tenant_id, access_token)

  return input1

def get_daily_emails(graph_api_endpoint,client_id, client_secret, redirect_uri, tenant_id, access_token):
    # Set up the headers with the access token
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    # Make the request to the Microsoft Graph API to retrieve emails for the specified user
    response = requests.get(graph_api_endpoint, headers=headers)
    # Check if the request was successful
    if response.status_code == 200:
        emails = response.json()
        # Save the entire response to a JSON file
        with open('data.json', 'w') as json_file:
            json.dump(emails, json_file, indent=4)
        for email in emails.get("value", []):
            print("------------------------------------------------------------------------------------------")
            #print(f"Subject: {email['subject']}")
            #print(f"From: {email['from']['emailAddress']['name']} ({email['from']['emailAddress']['address']})")
            print(f"Received: {email['receivedDateTime']}")
            print(f"Body: {email.get('body', {}).get('content')}")
            print("\n")
    else:
        print("Failed to retrieve emails. Status code:", response.status_code)
        print(response.text)


# Call the function to retrieve emails
my_python_tool("Hello World")