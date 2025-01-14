import logging
from datetime import datetime as dt
from datetime import time, timedelta
import os
import traceback
import tqdm
import app.exceptions as ex
import app.outlook as ol
import app.qualer_api as q

# Change the current working directory to the location of this file
new_directory = os.path.dirname(os.path.abspath(__file__))
os.chdir(new_directory)

##########################################################################################################
########################################### Initial Variables ############################################
##########################################################################################################

LIVE = False  # Set to True to run the script in live mode, False to run in test mode
# Set logging level to INFO
logging.basicConfig(filename='app/exception.log', level=logging.INFO)

# Initialize counters for success and failure
created_counter = 0
deleted_counter = 0
updated_counter = 0
skipped_counter = 0
failure_counter = 0

# Initialize an empty array to store exceptions
exceptions = []
created_events = []
deleted_events = []
updated_events = []
last_week = False
work_order_numbers = []

##########################################################################################################
########################################## Function Definitions ##########################################
##########################################################################################################


# Function that parses datetimes for combine_date_and_time()
def parse_datetime(datetime_str):
    return dt.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S")


# Function to combine date and time for an order. If dates are missing, raises an exception. If times are both missing, assumes all day event.
def combine_date_and_time(order):
    required_fields = ["RequestFromTime", "RequestToTime", "RequestFromDate", "RequestToDate"]
    # If no values are missing, assume the event is not all day
    if all(isinstance(order.get(field), str) for field in required_fields):
        request_from = dt.combine(parse_datetime(order["RequestFromDate"]).date(),
                                  parse_datetime(order["RequestFromTime"]).time())
        request_to = dt.combine(parse_datetime(order["RequestToDate"]).date(),
                                parse_datetime(order["RequestToTime"]).time())
        is_all_day = False
        # Check if the event ends before it starts
        if request_from > request_to:
            # If both times are AM, and the end time is before the start time, assume the end time is supposed to be PM
            if request_to.time() < time(12, 0):
                request_to += timedelta(hours=12)
                print("Order ends before it starts. Correcting end time from AM to PM on outlook.  Please make corrections on Qualer manually.")
                print("https://jgiquality.qualer.com/ServiceOrder/Info/" + str(order["ServiceOrderId"]))
            else:
                raise Exception("Order ends before it starts. Make manual corrections on Qualer: https://jgiquality.qualer.com/ServiceOrder/Info/" + order["ServiceOrderId"])

    # If the dates are present, check which of the times are missing
    elif all(isinstance(order.get(field), str) for field in required_fields[2:]):
        # If both times are missing, assume the event is all day
        if all(order.get(field) is None for field in required_fields[:2]):
            print(f"Order {order['ServiceOrderId']} has missing times. Assuming all day event.")
            request_from = dt.combine(parse_datetime(order["RequestFromDate"]).date(), time.min)
            request_to = dt.combine(parse_datetime(order["RequestToDate"]).date(), time.min) + timedelta(days=1)
            is_all_day = True
        # Use defualt times if only one of the times is missing
        else:
            default_start_time = time(7, 0)  # Default start time is 7:00 AM
            default_end_time = time(17, 0)  # Default end time is 5:00 PM

            start_time_str = order.get("RequestFromTime")
            end_time_str = order.get("RequestToTime")

            start_time = parse_datetime(start_time_str).time() if start_time_str else default_start_time
            end_time = parse_datetime(end_time_str).time() if end_time_str else default_end_time

            request_from = dt.combine(parse_datetime(order["RequestFromDate"]).date(), start_time)
            request_to = dt.combine(parse_datetime(order["RequestToDate"]).date(), end_time)

            is_all_day = False
    else:
        missing_values = [field for field in required_fields if not order.get(field)]
        raise Exception("Order is missing values: " + ", ".join(missing_values) + ".")

    return request_from, request_to, is_all_day


# Function to prepare event data as json for an order
def prepare_event_as_json(order):
    assignees = []  # Initialize list of assignees
    service_order_id = order["ServiceOrderId"]
    custom_order_number = order["CustomOrderNumber"]
    order_status = order["OrderStatus"]
    address = order["ShippingAddress"]
    address_str = f"{address['Address1']}, {address['City']}, {address['StateProvinceAbbreviation']} {address['ZipPostalCode']}"
    start_time, end_time, is_all_day = combine_date_and_time(order)
    order_assignments = q.get_work_order_assignments(service_order_id)
    for assignment in order_assignments:
        try:
            assignees.append(q.prepare_outlook_event_attendee(assignment["EmployeeId"]))
        except Exception as e:
            print(e)

    # Prepare body content for calendar event
    hyperlink = f'<a href="https://jgiquality.qualer.com/ServiceOrder/Info/{service_order_id}">{custom_order_number}</a>'
    body_content = "<b>" + hyperlink + "<br> Number of Assets:</b> " + str(q.count_assets(service_order_id))
    with open("app/body.html", 'r') as file:  # Read the contents of body.html
        body_html = file.read()
    body = body_html.replace('<p class="MsoNormal"></p>', '<p class="MsoNormal">' + body_content + '</p>')

    # Prepare dictionary object for Outlook calendar event
    event = {
        "subject": order["ClientCompanyName"],
        "bodyPreview": custom_order_number,
        "allowNewTimeProposals": False,
        "body": {
            "contentType": "html",
            "content": body
        },
        "isAllDay": is_all_day,
        "start": {
            "dateTime": start_time.strftime('%Y-%m-%dT%H:%M:%S.%f'),
            "timeZone": "America/Chicago"
        },
        "end": {
            "dateTime": end_time.strftime('%Y-%m-%dT%H:%M:%S.%f'),
            "timeZone": "America/Chicago"
        },
        "location": {
            "displayName": address_str,
            "locationType": "default"
        },
        "attendees": assignees,
        "categories": [],
        "showAs": "tentative" if order_status == "Scheduling" else "busy" if order_status == "Processing" else "free",
        "responseRequested": False,
        "isReminderOn": False,
        "isCancelled": True if order_status == "Cancelled" else False,
    }
    return event


def coerce_datetime_format(string_datetime):
    datetime_obj = dt.strptime(string_datetime, '%Y-%m-%dT%H:%M:%S.%f0')    # Coercing the string to a datetime object
    formatted_datetime = datetime_obj.strftime('%Y-%m-%dT%H:%M:%S.%f')      # Formatting the datetime object in the desired format
    return formatted_datetime


# Function to reformat an event to match the format of the event expected by the Outlook API
def reformat_event(event):
    reformatted_event = {}

    # Keys to compare between Qualer and Outlook, to determine whether or not to update an event
    keys_to_copy = [
        'subject',
        'bodyPreview',
        'allowNewTimeProposals',
        'isAllDay',
        'categories',
        'showAs',
        'responseRequested',
        'isReminderOn',
        'isCancelled'
    ]

    for key in keys_to_copy:
        reformatted_event[key] = event[key]

    reformatted_event['body'] = {
        'contentType': event['body']['contentType'],
        'content': event['body']['content']
    }

    reformatted_event['start'] = {
        'dateTime': coerce_datetime_format(event['start']['dateTime']),
        'timeZone': event['start']['timeZone']
    }

    reformatted_event['end'] = {
        'dateTime': coerce_datetime_format(event['end']['dateTime']),
        'timeZone': event['end']['timeZone']
    }

    reformatted_event['location'] = {
        'displayName': event['location']['displayName'],
        'locationType': event['location']['locationType']
    }

    reformatted_event['attendees'] = []
    for attendee in event['attendees']:
        # Remove ".onmicrosoft" from email addresses, just in case it's there
        attendee['emailAddress']['address'] = attendee['emailAddress']['address'].replace('.onmicrosoft', '')
        attendee_info = {
            'type': attendee['type'],
            'emailAddress': attendee['emailAddress']
        }
        reformatted_event['attendees'].append(attendee_info)

    return reformatted_event


# Function to find an event by its id from an outlook response
def find_event(event_id, all_outlook_events):
    for event in all_outlook_events['value']:
        if event['id'] == event_id:
            return event
    return None


# Function to compare two event objects and return a list of keys whose values differ between the two events
def compare_events(event1, event2):
    differing_keys = []
    for key in event1:
        if key == "body":  # or key == "bodyPreview":
            continue  # Do not check the body or bodyPreview keys; there are too many formatting discrepancies
        elif event1[key] != event2.get(key):
            differing_keys.append(key)
    return differing_keys


# Function to process an order
def process_order(order, id_array, is_live):
    event_id = ol.check_outlook_event(order["ServiceOrderId"], order["CustomOrderNumber"], id_array)
    if order.get("RequestToDate") is None:
        return None

    request_to_date = parse_datetime(order["RequestToDate"]).date()
    if request_to_date < dt.now().date():
        return "Past"  # Skip the order if it has passed the RequestToDate

    if order["OrderStatus"] == "Cancelled":
        if event_id:
            ol.delete_outlook_event(event_id) if is_live else print(f"Would have deleted event for {order['CustomOrderNumber']} if live")
            return "Cancelled"
        else:
            return "Skipped"  # Skip the cancelled order if it does not have an event in Outlook.

    qualer_event_obj = prepare_event_as_json(order)

    if event_id:
        outlook_event_obj = reformat_event(find_event(event_id, all_outlook_events))
        differing_keys = compare_events(outlook_event_obj, qualer_event_obj)
        if not differing_keys:
            print(f"Event for {order['CustomOrderNumber']} is up to date")
            return "Skipped"  # Skip if the changes to the order are irrelevant to the calendar event
        else:
            ol.update_outlook_event(event_id, qualer_event_obj, differing_keys == ['attendees']) if is_live else print(f"Would have updated event for {order['CustomOrderNumber']} if live")
            return "Updated"
    else:
        ol.create_outlook_event(qualer_event_obj) if is_live else print(f"Would have created event for {order['CustomOrderNumber']} if live")
        return "Created"


###########################################################################################
#################################### Main script ##########################################
###########################################################################################

# Get the last log time from the log file
last_log = ex.get_last_log_time()
if last_log is None:
    logging.critical("Last log time is None. Please check the log file.")
    raise SystemExit("Last log time is None. Please check the log file.")
try:
    last_log_datetime = dt.strptime(last_log, "%Y-%m-%d %H:%M:%S,%f")
except ValueError as e:
    logging.critical(f"Error parsing last log time: {e}")
    raise SystemExit("Error parsing last log time. Please check the log file.")

# Define the start and stop dates
start_date = last_log_datetime.replace(hour=0, minute=0, second=0, microsecond=0)
stop_date = dt.now()  # Use the current date as the stop date
print(start_date, stop_date)

# Get all Outlook events, and extract the event id's into a table
all_outlook_events = ol.get_outlook_events()
id_array = ol.extract_event_details(all_outlook_events)
print(f"Found {len(id_array)} events in Outlook")

# Construct the file path relative to the current directory
html_file_path = os.path.join(os.getcwd(), "app", "body.html")

# Open the file for reading or show an error message
try:
    with open(html_file_path, 'r') as file:
        body_html = file.read()
except FileNotFoundError:
    print(f"File {html_file_path} does not exist.")

# Set the initial week start and week end dates
week_start = start_date
week_end = start_date + timedelta(days=7)

pbar = tqdm.tqdm(total=(stop_date - start_date).days, desc="Processing Orders")
# Run the loop until the week end date is equal to or exceeds the stop date
while week_start <= stop_date:

    # Check to see if the week end date exceeds the stop date
    if week_end > stop_date:
        week_end = stop_date
        last_week = True

        iter_count = (week_end - week_start).days
    else:
        iter_count = 7
    pbar.update(iter_count)  # Update the progress bar

    # Call the get_work_orders function from ./app/qualer_api.py for the current week
    work_orders = q.get_work_orders(week_start, week_end)
    print(f"Found {len(work_orders)} records between {week_start.strftime('%Y-%m-%d')} and {week_end.strftime('%Y-%m-%d')}")

    # Loop through each work order
    for order in work_orders:
        try:
            event_id = ol.check_outlook_event(order["ServiceOrderId"], order["CustomOrderNumber"], id_array)
            result = process_order(order, id_array, LIVE)

            if result == "Past" or result == "Skipped":
                skipped_counter += 1
                continue
            elif result == "Cancelled":
                deleted_events.append(int(order["CustomOrderNumber"][6:]))
                deleted_counter += 1
            elif result == "Updated":
                updated_events.append(int(order["CustomOrderNumber"][6:]))
                updated_counter += 1
            elif result == "Created":
                created_events.append(int(order["CustomOrderNumber"][6:]))
                created_counter += 1

        except ValueError:
            failure_counter += 1
            exceptions.append([order["CustomOrderNumber"][6:], traceback.format_exc()])
        except Exception as e:
            failure_counter += 1
            exceptions.append([order["CustomOrderNumber"][6:], str(e)])

    if (last_week):
        break

    # Update the week start and week end dates for the next iteration
    week_start = week_end
    week_end = week_start + timedelta(days=7)


###########################################################################################
###################################### Logging ############################################
###########################################################################################

# Logging
ex.group_orders_by_exception(exceptions)  # log the exceptions
logging.info(f"Successfully created orders: {created_events}") if created_events else None
logging.info(f"Successfully updated orders: {updated_events}") if updated_events else None
logging.info(f"Successfully deleted orders: {deleted_events}") if deleted_events else None


print()

# Print the summary of CustomOrderNumbers count per unique exception
for exception, count in ex.count_exceptions(exceptions).items():
    print(f"Exception: {exception} (Count: {count})")
print()

# Print the results of the script
print()
print(f"Created: {created_counter}")
print(f"Updated: {updated_counter}")
print(f"Deleted: {deleted_counter}")
print(f"Skipped: {skipped_counter}")
print(f"Failed: {failure_counter}")
print()
