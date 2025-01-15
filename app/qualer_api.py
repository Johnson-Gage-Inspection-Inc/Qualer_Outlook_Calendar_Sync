import requests
import json
import time
from os import environ


class QualerAPI:
    def __init__(self, GENERATE_TOKEN=True):
        """_summary_

        Args:
            GENERATE_TOKEN (bool, optional): While testing, set to true for first time use, then set to false and replace QUALER_API_KEY with the generated token. Defaults to True.
        """

        self.endpoint = "https://jgiquality.qualer.com/api"
        if GENERATE_TOKEN:
            self.generate_token()
        else:
            self.token = "Api-Token fa0a2ee8-fb72-463c-8fcb-da9cfd4861e0"  # Replace with the current API token, if desired
        self.headers = {
            'Content-Type': 'application/json',
            'Authorization': self.token
        }

    def login(self, username, password):
        """Function to log in to Qualer API and retrieve a token"""
        endpoint = f"{self.endpoint}/login"
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

    def generate_token(self):
        """Function to generate a fresh Qualer API token"""
        user = environ.get('QUALER_USER')
        pw = environ.get('QUALER_PASSWORD')
        self.token = "Api-Token " + self.login(user, pw)
        # print(f"Qualer API Token: {self.token}")
        return True

    def qualer_error_handler(self, response):
        """Function to handle Qualer API response errors"""
        if response.status_code == 200:
            return response.json()
        elif response.status_code == 400:
            print(f"400 Error: Bad request. {response.text}")
            raise Exception(f"400 Error: Bad request. {response.text}")
        elif response.status_code == 401:
            print("401 Error: Generating new token...")
            self.generate_token()  # Generate new Qualer API token
            return None
        elif response.status_code == 404:
            print(f"404 Error: No data found at {self.endpoint}. {response.text}")
            raise Exception(f"404 Error: No data found at {self.endpoint}. {response.text}")
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

    def get_work_orders(self, start, end):
        """Function to retrieve future work orders from Qualer API"""
        while True:
            start_time = start.strftime('%Y-%m-%dT%H:%M:%S.%f')
            end_time = end.strftime('%Y-%m-%dT%H:%M:%S.%f')
            response = requests.get(self.endpoint + f"/service/workorders?status=OnSite&from={start_time}&to={end_time}", headers=self.headers)
            response.raise_for_status()
            work_orders = response.json()
            if work_orders is not None:
                self.work_orders = work_orders
                return

    # def get_work_order(self, workOrderNumber):
    #     """Function to retrieve work order details from Qualer API"""
    #     while True:
    #         response = requests.get(self.endpoint + f"/service/workorders?workOrderNumber={workOrderNumber}", headers=self.headers)
    #         if self.qualer_error_handler(response) is not None:
    #             return response.json()

    def count_assets(self, serviceOrderId):
        """Function to count assets on an order"""
        while True:
            response = requests.get(self.endpoint + f"/service/workorders/{serviceOrderId}/workitems", headers=self.headers)
            if self.qualer_error_handler(response) is not None:
                return len(response.json())

    def get_work_order_assignments(self, serviceOrderId):
        """Function to retrieve assignments for an order from Qualer API"""
        while True:
            response = requests.get(self.endpoint + f"/service/workorders/{serviceOrderId}/assignments", headers=self.headers)
            if self.qualer_error_handler(response) is not None:
                return response.json()

    def prepare_event_attendee(self, employee):
        """Function to look up an employee and transform the data into a format suitable for an Outlook calendar event attendee"""
        EmployeeId = employee['EmployeeId']
        while True:
            response = requests.get(self.endpoint + f"/employees/{EmployeeId}", headers=self.headers)
            if self.qualer_error_handler(response) is not None:
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
                return self.qualer_error_handler(response)
