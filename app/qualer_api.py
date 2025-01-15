import requests
import json
import time
import os

LIVE = False  # Set to True to run the script in live mode, False to run in test mode

###########################################################################################################
############################################ Qualer API Calls #############################################
###########################################################################################################


# Function to log in to Qualer API and retrieve a token
def login(endpoint, username, password):
    endpoint = f"{endpoint}/login"
    header = {
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    data = {
        "UserName": username,
        "Password": password,
        "ClearPreviousTokens": "False"
    }

    try:
        response = requests.post(endpoint, data=json.dumps(data), headers=header)
        response.raise_for_status()
        return response.json()['Token']
    except (KeyError, ValueError) as e:
        print(f"Qualer API Login Exception: {e}")
        print(f"RESPONSE: {response.text}")
        raise


# Function to generate Qualer API token
def generate_token():
    global QUALER_API_KEY
    QUALER_API_KEY = "Api-Token " + login(QUALER_API_ENDPOINT, LOGIN_USER, LOGIN_PASS)
    print(f"Qualer API Token: {QUALER_API_KEY}")
    return True


# Qualer API configuration
LOGIN_USER = os.environ.get('QUALER_USER')
LOGIN_PASS = os.environ.get('QUALER_PASSWORD')
QUALER_API_ENDPOINT = "https://jgiquality.qualer.com/api"
GENERATE_TOKEN = True  # While testing, set to true for first time use, then set to false and replace QUALER_API_KEY with the generated token
if GENERATE_TOKEN:
    generate_token()  # Generate Qualer API token
else:
    QUALER_API_KEY = "Api-Token fa0a2ee8-fb72-463c-8fcb-da9cfd4861e0"  # Replace with the current API token, if desired
QUALER_API_HEADERS = {
    'Content-Type': 'application/json',
    'Authorization': f'{QUALER_API_KEY}'
}


# Function to handle Qualer API response errors
def qualer_error_handler(response):
    if response.status_code == 200:
        return response.json()
    elif response.status_code == 400:
        print(f"400 Error: Bad request. {response.text}")
        raise Exception(f"400 Error: Bad request. {response.text}")
    elif response.status_code == 401:
        print("401 Error: Generating new token...")
        generate_token()  # Generate new Qualer API token
        return None
    elif response.status_code == 404:
        print(f"404 Error: No data found at {QUALER_API_ENDPOINT}. {response.text}")
        raise Exception(f"404 Error: No data found at {QUALER_API_ENDPOINT}. {response.text}")
    elif response.status_code == 429:
        print("429 Error: Too many requests. Waiting 30 seconds...")
        time.sleep(30)  # Wait 30 seconds for Qualer API to become available
        return None
    elif response.status_code == 503:
        print("503 Error: Qualer API unavailable. Waiting 30 seconds...")
        time.sleep(30)  # Wait 30 seconds for Qualer API to become available
        return None
    else:
        raise Exception(f"Qualer API error: {response.status_code} {response.text}")


# Function to retrieve future work orders from Qualer API
def get_work_orders(start, end):
    while True:
        start_time = start.strftime('%Y-%m-%dT%H:%M:%S.%f')
        end_time = end.strftime('%Y-%m-%dT%H:%M:%S.%f')
        response = requests.get(QUALER_API_ENDPOINT + f"/service/workorders?status=OnSite&from={start_time}&to={end_time}", headers=QUALER_API_HEADERS)
        work_orders = qualer_error_handler(response)
        if work_orders is not None:
            return work_orders


# Function to retrieve work order details from Qualer API
def get_work_order(workOrderNumber):
    while True:
        response = requests.get(QUALER_API_ENDPOINT + f"/service/workorders?workOrderNumber={workOrderNumber}", headers=QUALER_API_HEADERS)
        if qualer_error_handler(response) is not None:
            return response.json()


# Function to count assets on an order
def count_assets(serviceOrderId):
    while True:
        response = requests.get(QUALER_API_ENDPOINT + f"/service/workorders/{serviceOrderId}/workitems", headers=QUALER_API_HEADERS)
        if qualer_error_handler(response) is not None:
            return len(response.json())


# Function to retrieve assignments for an order from Qualer API
def get_work_order_assignments(serviceOrderId):
    while True:
        response = requests.get(QUALER_API_ENDPOINT + f"/service/workorders/{serviceOrderId}/assignments", headers=QUALER_API_HEADERS)
        if qualer_error_handler(response) is not None:
            return response.json()


# Function to look up an employee and transform the data into a format suitable for an Outlook calendar event attendee
def prepare_outlook_event_attendee(EmployeeId):
    while True:
        response = requests.get(QUALER_API_ENDPOINT + f"/employees/{EmployeeId}", headers=QUALER_API_HEADERS)
        if qualer_error_handler(response) is not None:
            api_response = response.json()  # Extract the API response as JSON data

            # Transform the employee information into a format suitable for an Outlook calendar event attendee
            transformed_data = {
                "type": "required",
                "emailAddress": {
                    "name": f"{api_response['FirstName']} {api_response['LastName']}",
                    "address": api_response['SubscriptionEmail']
                }
            }
            return transformed_data
        else:
            return qualer_error_handler(response)
