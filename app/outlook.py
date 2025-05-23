import requests
from bs4 import BeautifulSoup
import re
import json
import traceback

###########################################################################################################
############################################### Outlook API ###############################################
###########################################################################################################


# Outlook API authorization
def get_access_token():
    try:
        tenant_id = "9def3ae4-854a-4465-952c-5693835965d9"
        token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        payload = {
            "grant_type": "client_credentials",
            "client_id": "7d2eae58-703b-418d-8291-17e4f9c53a40",
            "client_secret": "M3f8Q~Fmr5oWhMJsdEkL~n6F6vu4_L9b~OP3Xa28",
            "scope": "https://graph.microsoft.com/.default"
        }
        response = requests.post(token_url, data=payload)
        access_token = response.json()["access_token"]
        return access_token
    except Exception as e:
        print(f"Outlook API Authorization Exception: {e}")
        raise


access_token = get_access_token()
calendar_id = 'AAMkAGEwOWUyZDEzLTQ1MTktNDNkMy1hZmZiLTQxZjZmNGVmNGZlMABGAAAAAACxzJLn1GFkQpJCvD31IsIGBwAxNg00JTgmTqbPLRQZN89GAAAAAAEHAAAxNg00JTgmTqbPLRQZN89GAAjGS-TTAAA='
user_id = 'sysop@jgiquality.com'
endpoint = f'https://graph.microsoft.com/v1.0/users/{user_id}/'
headers = {
    'Authorization': 'Bearer ' + access_token,
    'Prefer': 'outlook.timezone = "America/Chicago"',
    'Content-Type': 'application/json'
}


# Function to format an error message from outlook API
def outlook_error_handler(error):
    error_code = error['code']
    error_message = error['message']
    formatted_error = f"Error: {error_code}\nMessage: {error_message}"
    return formatted_error


# Function to create an array of IDs for all Outlook events
def extract_event_details(all_outlook_events):
    existing_events = []
    pattern = r"56561-\d{6}"  # Regex pattern for "56561-" followed by 6 digits

    for event in all_outlook_events['value']:
        event_id = event['id']
        body_preview = event['bodyPreview']
        custom_order_number = re.search(pattern, body_preview)
        if custom_order_number:
            custom_order_number = custom_order_number.group(0)
        else:
            custom_order_number = None

        try:  # Try to get the ServiceOrderId from the HTML body, by looking for the first <a> tag with an href starting with the desired URL
            html = event['body']['content']
            soup = BeautifulSoup(html, 'html.parser')  # Parse the HTML using BeautifulSoup
            matching_a_tags = soup.find_all('a', href=lambda href: href.startswith('https://jgiquality.qualer.com/ServiceOrder/Info/'))  # Find all <a> tags with href starting with the desired URL
            service_order_id = [a_tag['href'].split('/')[-1] for a_tag in matching_a_tags][0]  # Extract the values from the href attributes, and assume the first one is the service order ID
        except Exception:
            service_order_id = None
        finally:
            existing_events.append([service_order_id, custom_order_number, event_id])

    return existing_events


# Function to check if an Outlook event exists
def check_outlook_event(ServiceOrderId, CustomOrderNumber, id_array):
    event_id = False
    for event in id_array:
        if ServiceOrderId == event[0] or CustomOrderNumber == event[1]:
            event_id = event[2]
            break
    return event_id

#######################################################################################
##################################  CRUD Operations  ##################################
#######################################################################################


# C: Function to create an Outlook calendar event
def create_outlook_event(event):
    url = f'{endpoint}calendars/{calendar_id}/events'
    data = event
    response = requests.post(url, headers=headers, json=data)

    if response.status_code == 201:
        print('Event created successfully for ' + event['bodyPreview'] + '.')
    elif response.status_code == 400:
        raise Exception(outlook_error_handler(response.json().get('error', {})))
    else:
        print('Failed to create event for ' + event['bodyPreview'] + '.')
        print(response.content)
        raise Exception("Create error: ", str(response))
    return response.json()


# R: Function to retrieve Outlook calendar events
def get_outlook_events():
    global access_token
    first_attempt = True  # Flag to indicate if this is the first attempt to get the events
    url = f'{endpoint}calendars/{calendar_id}/events'  # Endpoint to get the events
    events = {}  # Initialize the events dictionary
    skip = 0  # Initial value for $skip
    top = 1000  # Maximum value for $top (max 1000 events per request)

    while True:
        params = {'$top': top, '$skip': skip}
        response = requests.get(url, headers=headers, params=params)

        try:
            response.raise_for_status()
        except requests.exceptions.HTTPError as e:
            if response.status_code == 401 and first_attempt:
                access_token = get_access_token()
                first_attempt = False
                continue
            else:
                print(f"HTTPError: {e}")
                print("Response content:", response.content)
                traceback.print_exc()
                raise

        try:
            data = response.json()
        except json.decoder.JSONDecodeError:
            print("JSONDecodeError: ", response.content)
            traceback.print_exc()
            raise

        if not data:
            if not first_attempt:
                raise Exception(outlook_error_handler(data.get('error', {})))
            else:
                first_attempt = False
                continue

        events['value'] = events.get('value', []) + data['value']
        if len(data['value']) < top:
            break

        skip += top

    return events


# U: Function to update an Outlook event
def update_outlook_event(event_id, event, attendees_only=False):
    url = f"{endpoint}events/{event_id}"
    request_body = json.loads('{"attendees":' + json.dumps(event['attendees']) + '}') if attendees_only else event
    response = requests.patch(url, headers=headers, data=json.dumps(request_body))

    # Check the response status
    if response.status_code == 200:
        print("Event updated successfully for " + event['bodyPreview'] + ".")
    else:
        print('Failed to update event for ' + event['bodyPreview'] + '.')
        raise Exception("Update error: ", response.text)
    return response.json()


# D: Function to delete an Outlook event
def delete_outlook_event(event_id):
    url = f"{endpoint}events/{event_id}"
    response = requests.delete(url, headers=headers)

    # Check the response status
    if response.status_code == 204:
        print("Event deleted successfully.")
        return
    else:
        print('Failed to delete event.')
        print(response.content)
        raise Exception("Delete error: ", response.text)
