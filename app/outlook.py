import requests
from bs4 import BeautifulSoup
import re
import json
import traceback
import logging
from datetime import datetime as dt


class OLEvent(dict):
    def __init__(self, event):
        self.update({
            'subject': event['subject'],
            'bodyPreview': event['bodyPreview'],
            'allowNewTimeProposals': event['allowNewTimeProposals'],
            'isAllDay': event['isAllDay'],
            'categories': event['categories'],
            'showAs': event['showAs'],
            'responseRequested': event['responseRequested'],
            'isReminderOn': event['isReminderOn'],
            'isCancelled': event['isCancelled'],
            'body': {
                'contentType': event['body']['contentType'],
                'content': event['body']['content']
            },
            'start': {
                'dateTime': Outlook.coerce_datetime_format(event['start']['dateTime']),
                'timeZone': event['start']['timeZone']
            },
            'end': {
                'dateTime': Outlook.coerce_datetime_format(event['end']['dateTime']),
                'timeZone': event['end']['timeZone']
            },
            'location': {
                'displayName': event['location']['displayName'],
                'locationType': event['location']['locationType']
            },
            'attendees': [self.formatAttendee(attendee) for attendee in event['attendees']]
        })

    def formatAttendee(self, attendee):
        """Function to format an attendee object from an outlook response

        Args:
            self (_type_): _description_
            attendee (_type_): _description_

        Returns:
            _type_: _description_
        """
        email = attendee['emailAddress']
        email['address'] = email['address'].replace('.onmicrosoft', '')
        return {
            'type': attendee['type'],
            'emailAddress': email
        }


class Outlook:
    def __init__(self, user_id='sysop@jgiquality.com'):
        self.get_access_token()
        self.calendar_id = 'AAMkAGEwOWUyZDEzLTQ1MTktNDNkMy1hZmZiLTQxZjZmNGVmNGZlMABGAAAAAACxzJLn1GFkQpJCvD31IsIGBwAxNg00JTgmTqbPLRQZN89GAAAAAAEHAAAxNg00JTgmTqbPLRQZN89GAAjGS-TTAAA='
        self.endpoint = f'https://graph.microsoft.com/v1.0/users/{user_id}/'
        self.headers = {
            'Authorization': self.token,
            'Prefer': 'outlook.timezone = "America/Chicago"',
            'Content-Type': 'application/json'
        }
        self.events = {}
        self.event = self.event(self)

    def get_access_token(self, tenant_id="9def3ae4-854a-4465-952c-5693835965d9"):
        try:
            token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
            payload = {
                "grant_type": "client_credentials",
                "client_id": "7d2eae58-703b-418d-8291-17e4f9c53a40",
                "client_secret": "M3f8Q~Fmr5oWhMJsdEkL~n6F6vu4_L9b~OP3Xa28",
                "scope": "https://graph.microsoft.com/.default"
            }
            response = requests.post(token_url, data=payload)
            response.raise_for_status()
            token = response.json()["access_token"]
            self.token = f'Bearer {token}'
        except Exception as e:
            logging.critical(f"Outlook API Authorization Exception: {e}")
            raise

    def error_handler(error):
        """Function to format an error message from outlook API"""
        error_code = error['code']
        error_message = error['message']
        formatted_error = f"Error: {error_code}\nMessage: {error_message}"
        return formatted_error

    def getEvents(self):
        """Function to create an array of IDs for all Outlook events"""
        existing_events = []
        pattern = r"56561-\d{6}"  # Regex pattern for "56561-" followed by 6 digits

        for event in self.events['value']:
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

        self.id_array = existing_events
        logging.info(f"Found {len(self.id_array)} events in Outlook")

    def check(self, ServiceOrderId, CustomOrderNumber):
        """Function to check if an event already exists in Outlook"""
        self.event_id = False
        for event in self.id_array:
            if ServiceOrderId == event[0] or CustomOrderNumber == event[1]:
                self.event_id = event[2]
                break
        return self.event_id

    @staticmethod
    def coerce_datetime_format(string_datetime):
        datetime_obj = dt.strptime(string_datetime, '%Y-%m-%dT%H:%M:%S.%f0')    # Coercing the string to a datetime object
        formatted_datetime = datetime_obj.strftime('%Y-%m-%dT%H:%M:%S.%f')      # Formatting the datetime object in the desired format
        return formatted_datetime

    class event:
        def __init__(self, parent: 'Outlook'):
            self.parent = parent
            self.read()

        def lookup(self, event_id) -> OLEvent:
            """Function to find an event by its id from an outlook response"""
            for event in self.parent.events['value']:
                if event['id'] == event_id:
                    return OLEvent(event)

        # C:
        def create(self, event: dict):
            """Function to create an Outlook calendar event"""
            url = f'{self.parent.endpoint}calendars/{self.parent.calendar_id}/events'
            response = requests.post(url, headers=self.parent.headers, json=event)

            if response.status_code == 201:
                print('Event created successfully for ' + event['bodyPreview'] + '.')
            elif response.status_code == 400:
                raise Exception(self.parent.error_handler(response.json().get('error', {})))
            else:
                print('Failed to create event for ' + event['bodyPreview'] + '.')
                print(response.content)
                raise Exception("Create error: ", str(response))
            return response.json()

        # R:
        def read(self, skip=0, top=1000):
            """R: Get all Outlook events, and extract the event id's into a table.

            Args:
                skip (int, optional): Initial value for $skip. Defaults to 0.
                top (int, optional): Maximum value for $top (max 1000 events per request). Defaults to 1000.

            Raises:
                Exception: _description_
            """
            first_attempt = True  # Flag to indicate if this is the first attempt to get the events
            url = f'{self.parent.endpoint}calendars/{self.parent.calendar_id}/events'  # Endpoint to get the events

            print("Getting events from outlook..", end='')
            while True:
                print('.', end='')
                params = {'$top': top, '$skip': skip}
                response = requests.get(url, headers=self.parent.headers, params=params)

                try:
                    response.raise_for_status()
                except requests.exceptions.HTTPError as e:
                    if response.status_code == 401 and first_attempt:
                        self.parent.get_access_token()
                        first_attempt = False
                        continue
                    else:
                        logging.debug(f"HTTPError: {e}")
                        logging.debug("Response content:", response.content)
                        traceback.print_exc()
                        raise

                if not (data := response.json()):
                    if first_attempt:
                        continue
                    raise Exception(self.parent.error_handler(data.get('error', {})))

                self.parent.events['value'] = self.parent.events.get('value', []) + data['value']
                if len(data['value']) < top:
                    break

                skip += top
            self.parent.getEvents()

        # U:
        def update(self, event_id, event: dict, attendees_only=False):
            """Function to update an Outlook event"""
            url = f"{self.parent.endpoint}events/{event_id}"
            request_body = json.loads('{"attendees":' + json.dumps(event['attendees']) + '}') if attendees_only else event
            response = requests.patch(url, headers=self.parent.headers, data=json.dumps(request_body))
            response.raise_for_status()
            if response.status_code == 200:
                logging.info("Event updated successfully for " + event['bodyPreview'] + ".")
                return response.json()
            else:
                print('Failed to update event for ' + event['bodyPreview'] + '.')
                raise Exception("Update error: ", response.text)

        # D:
        def delete(self, event_id):
            """Function to delete an Outlook event"""
            url = f"{self.parent.endpoint}events/{event_id}"
            response = requests.delete(url, headers=self.parent.headers)
            response.raise_for_status()
            if response.status_code == 204:
                print("Event deleted successfully.")
                return
