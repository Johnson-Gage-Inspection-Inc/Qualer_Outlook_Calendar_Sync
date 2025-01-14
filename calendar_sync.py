import logging
from datetime import datetime as dt
from datetime import time, timedelta
import os
import traceback
from tqdm import tqdm
import app.exceptions as ex
import app.outlook as ol
import app.qualer_api as q

# Set the current directory to the directory of the script
new_directory = os.path.dirname(os.path.abspath(__file__))
os.chdir(new_directory)

# Set up logging
logging.basicConfig(filename='app/exception.log', level=logging.DEBUG)


class CalendarSync:
    def __init__(self):
        self.created_counter = 0
        self.deleted_counter = 0
        self.updated_counter = 0
        self.skipped_counter = 0
        self.failure_counter = 0
        self.exceptions = []
        self.created_events = []
        self.deleted_events = []
        self.updated_events = []
        self.last_week = False
        self.work_order_numbers = []

    def loopOrders(cSync, work_orders):
        """Function to loop through work orders and process them

        Args:
            cSync (CalendarSync): The core CalendarSync object
            work_orders (list): A list of work orders
            LIVE (bool, optional): Set to True to run the script in live mode, False to run in test mode. Defaults to False.
        """
        for work_order in tqdm(work_orders, desc="Processing Orders", leave=False, dynamic_ncols=True):
            try:
                CustomOrderNumber = work_order["CustomOrderNumber"][6:]
                order = QualerOrder(work_order)
                order.status = order.process_order(cSync.id_array)
                if cSync.tickCounters(order):
                    continue

            except ValueError:
                cSync.failure_counter += 1
                cSync.exceptions.append([CustomOrderNumber, traceback.format_exc()])
            except Exception as e:
                cSync.failure_counter += 1
                cSync.exceptions.append([CustomOrderNumber, str(e)])

    def tickCounters(cSync, order):
        if order.status == "Past" or order.status == "Skipped":
            cSync.skipped_counter += 1
            return True
        elif order.status == "Cancelled":
            cSync.deleted_events.append(order.CustomOrderNumber)
            cSync.deleted_counter += 1
        elif order.status == "Updated":
            cSync.updated_events.append(order.CustomOrderNumber)
            cSync.updated_counter += 1
        elif order.status == "Created":
            cSync.created_events.append(order.CustomOrderNumber)
            cSync.created_counter += 1

    def finalLogging(cSync):
        ex.group_orders_by_exception(cSync.exceptions)  # log the exceptions
        logging.info(f"Successfully created orders: {cSync.created_events}") if cSync.created_events else None
        logging.info(f"Successfully updated orders: {cSync.updated_events}") if cSync.updated_events else None
        logging.info(f"Successfully deleted orders: {cSync.deleted_events}") if cSync.deleted_events else None

        # Print the summary of CustomOrderNumbers count per unique exception
        for exception, count in ex.count_exceptions(cSync.exceptions).items():
            logging.info(f"Exception: {exception} (Count: {count})")


class QualerOrder(dict):
    def __init__(order, *args, **kwargs):
        super(QualerOrder, order).__init__(*args, **kwargs)
        requiredKeys = ["ServiceOrderId", "CustomOrderNumber", "OrderStatus"]  # TODO: Add more required keys
        assert all(key in order for key in requiredKeys), f"Missing required keys: {', '.join([key for key in requiredKeys if key not in order])}"
        order.id = order["ServiceOrderId"]
        order.number = order["CustomOrderNumber"]
        order.status = order["OrderStatus"]

        order.start_time = None
        order.end_time = None
        order.is_all_day = False
        order.combine_date_and_time()

    def combine_date_and_time(order):
        """Function to combine date and time for an order. If dates are missing, raises an exception. If times are both missing, assumes all day event."""
        required_fields = ["RequestFromTime", "RequestToTime", "RequestFromDate", "RequestToDate"]
        # If no values are missing, assume the event is not all day
        if all(isinstance(order.get(field), str) for field in required_fields):
            order.start_time = dt.combine(parse_datetime(order["RequestFromDate"]).date(),
                                          parse_datetime(order["RequestFromTime"]).time())
            order.end_time = dt.combine(parse_datetime(order["RequestToDate"]).date(),
                                        parse_datetime(order["RequestToTime"]).time())

            if order.start_time > order.end_time:
                if order.end_time.time() < time(12, 0):
                    order.end_time += timedelta(hours=12)
                    logging.warning("Order ends before it starts. Correcting end time from AM to PM on outlook.  Please make corrections on Qualer manually.")
                    logging.info("https://jgiquality.qualer.com/ServiceOrder/Info/" + str(order["ServiceOrderId"]))
                else:
                    raise Exception("Order ends before it starts. Make manual corrections on Qualer: https://jgiquality.qualer.com/ServiceOrder/Info/" + order["ServiceOrderId"])

        # If the dates are present, check which of the times are missing
        elif any(isinstance(order.get(field), str) for field in required_fields[2:]):
            # If one date is missing but not the other, assume they're the same.
            if not all(isinstance(order.get(field), str) for field in required_fields[2:]):
                # Assume the start and end dates are the same:
                if order.get("RequestFromDate") and not order.get("RequestToDate"):
                    order["RequestToDate"] = order["RequestFromDate"]
                elif not order.get("RequestFromDate") and order["RequestToDate"]:
                    order["RequestFromDate"] = order["RequestToDate"]
            # If both times are missing, assume the event is all day
            if all(order.get(field) is None for field in required_fields[:2]):
                logging.warning(f"Order {order['ServiceOrderId']} has missing times. Assuming all day event.")
                order.start_time = dt.combine(parse_datetime(order["RequestFromDate"]).date(), time.min)
                order.end_time = dt.combine(parse_datetime(order["RequestToDate"]).date(), time.min) + timedelta(days=1)
                order.is_all_day = True
            # Use defualt times if only one of the times is missing
            else:
                default_start_time = time(7, 0)  # Default start time is 7:00 AM
                default_end_time = time(17, 0)  # Default end time is 5:00 PM

                start_time_str = order.get("RequestFromTime")
                end_time_str = order.get("RequestToTime")

                start_time = parse_datetime(start_time_str).time() if start_time_str else default_start_time
                end_time = parse_datetime(end_time_str).time() if end_time_str else default_end_time

                order.start_time = dt.combine(parse_datetime(order["RequestFromDate"]).date(), start_time)
                order.end_time = dt.combine(parse_datetime(order["RequestToDate"]).date(), end_time)
        else:
            missing_values = [field for field in required_fields if not order.get(field)]
            raise Exception("Order is missing values: " + ", ".join(missing_values) + ".")

    def prepare_event_as_json(order) -> dict:
        """Function to map a Qualer order to an Outlook event JSON"""
        return {
            "subject": order["ClientCompanyName"],
            "bodyPreview": order.number,
            "allowNewTimeProposals": False,
            "body": {
                "contentType": "html",
                "content": order.body()
            },
            "isAllDay": order.is_all_day,
            "start": {
                "dateTime": order.start_time.strftime('%Y-%m-%dT%H:%M:%S.%f'),
                "timeZone": "America/Chicago"
            },
            "end": {
                "dateTime": order.end_time.strftime('%Y-%m-%dT%H:%M:%S.%f'),
                "timeZone": "America/Chicago"
            },
            "location": {
                "displayName": order.extractAddressStr(),
                "locationType": "default"
            },
            "attendees": order.gatherAssignees(),
            "categories": [],
            "showAs": "tentative" if order.status == "Scheduling" else "busy" if order.status == "Processing" else "free",
            "responseRequested": False,
            "isReminderOn": False,
            "isCancelled": True if order.status == "Cancelled" else False,
        }

    def body(order):
        hyperlink = f'<a href="https://jgiquality.qualer.com/ServiceOrder/Info/{order.id}">{order.number}</a>'
        body_content = f"<b>{hyperlink}<br> Number of Assets:</b> {q.count_assets(order.id)}"
        with open("app/body.html", 'r') as file:  # Read the contents of body.html
            body_html = file.read()
        body = body_html.replace('<p class="MsoNormal"></p>', '<p class="MsoNormal">' + body_content + '</p>')
        return body

    def extractAddressStr(order):
        """Function to extract the address string for an order

        Args:
            order (QualerOrder): The order dictionary
        """
        address = order["ShippingAddress"]
        return f"{address['Address1']}, {address['City']}, {address['StateProvinceAbbreviation']} {address['ZipPostalCode']}"

    def gatherAssignees(order):
        assignees = []
        order.assignments = q.get_work_order_assignments(order.id)
        for assignment in order.assignments:
            try:
                assignees.append(q.prepare_outlook_event_attendee(assignment["EmployeeId"]))
            except Exception as e:
                logging.error(e)
        return assignees

    def process_order(order, id_array, is_live=False):
        """Function to process an order"""
        order.CustomOrderNumber = int(order["CustomOrderNumber"][6:])
        event_id = ol.check_outlook_event(order["ServiceOrderId"], order["CustomOrderNumber"], id_array)
        if order.get("RequestToDate") is None:
            return None

        request_to_date = parse_datetime(order["RequestToDate"]).date()
        if request_to_date < dt.now().date():
            return "Past"  # Skip the order if it has passed the RequestToDate

        if order["OrderStatus"] == "Cancelled":
            if event_id:
                ol.delete_outlook_event(event_id) if is_live else logging.info(f"Would have deleted event for {order['CustomOrderNumber']} if live")
                return "Cancelled"
            else:
                return "Skipped"  # Skip the cancelled order if it does not have an event in Outlook.

        qualer_event_obj = order.prepare_event_as_json()

        if event_id:
            outlook_event_obj = reformat_event(find_event(event_id, all_outlook_events))
            differing_keys = compare_events(outlook_event_obj, qualer_event_obj)
            if not differing_keys:
                logging.info(f"Event for {order['CustomOrderNumber']} is up to date")
                return "Skipped"  # Skip if the changes to the order are irrelevant to the calendar event
            else:
                ol.update_outlook_event(event_id, qualer_event_obj, differing_keys == ['attendees']) if is_live else logging.info(f"Would have updated event for {order['CustomOrderNumber']} if live")
                return "Updated"
        else:
            ol.create_outlook_event(qualer_event_obj) if is_live else logging.info(f"Would have created event for {order['CustomOrderNumber']} if live")
            return "Created"


def parse_datetime(datetime_str):
    """Function that parses datetimes for combine_date_and_time()"""
    return dt.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S")


def coerce_datetime_format(string_datetime):
    datetime_obj = dt.strptime(string_datetime, '%Y-%m-%dT%H:%M:%S.%f0')    # Coercing the string to a datetime object
    formatted_datetime = datetime_obj.strftime('%Y-%m-%dT%H:%M:%S.%f')      # Formatting the datetime object in the desired format
    return formatted_datetime


def reformat_event(event):
    """Function to reformat an event to match the format of the event expected by the Outlook API"""
    reformatted_event = {
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
            'dateTime': coerce_datetime_format(event['start']['dateTime']),
            'timeZone': event['start']['timeZone']
        },
        'end': {
            'dateTime': coerce_datetime_format(event['end']['dateTime']),
            'timeZone': event['end']['timeZone']
        },
        'location': {
            'displayName': event['location']['displayName'],
            'locationType': event['location']['locationType']
        },
        'attendees': []

    }

    for attendee in event['attendees']:
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


if __name__ == "__main__":
    cSync = CalendarSync()
    last_log = ex.get_last_log_time()  # Get the last log time from the log file
    if not last_log:
        # TODO: Set up a way to notify myself when the last log time is None (email)
        raise SystemExit("Last log time is None. Please check the log file.")
    try:
        last_log_datetime = dt.strptime(last_log, "%Y-%m-%d %H:%M:%S,%f")
    except ValueError as e:
        logging.critical(f"Error parsing last log time: {e}")
        raise SystemExit("Error parsing last log time. Please check the log file.")

    # Define the start and stop dates
    start_date = last_log_datetime.replace(hour=0, minute=0, second=0, microsecond=0)
    stop_date = dt.now()  # Use the current date as the stop date

    # Get all Outlook events, and extract the event id's into a table
    all_outlook_events = ol.get_outlook_events()
    cSync.id_array = ol.extract_event_details(all_outlook_events)
    logging.info(f"Found {len(cSync.id_array)} events in Outlook")

    try:
        html_file_path = os.path.join(os.getcwd(), "app", "body.html")
        with open(html_file_path, 'r') as file:
            body_html = file.read()
    except FileNotFoundError:
        logging.critical("File /app/body.html does not exist.")

    # Set the initial week start and week end dates
    week_start = start_date
    week_end = start_date + timedelta(days=7)

    if (weeks := round((stop_date - start_date).days / 7, 1)) > 0:
        pbar = tqdm(total=weeks, desc="Iterating weeks", unit="weeks", dynamic_ncols=True)
    # Run the loop until the week end date is equal to or exceeds the stop date

    last_week = False
    while week_start <= stop_date:
        if week_end > stop_date:
            week_end = stop_date
            last_week = True
        if weeks > 0:
            pbar.update(1)  # Update the progress bar

        # Call the get_work_orders function from ./app/qualer_api.py for the current week
        work_orders = q.get_work_orders(week_start, week_end)
        logging.info(f"Found {len(work_orders)} records between {week_start.strftime('%Y-%m-%d')} and {week_end.strftime('%Y-%m-%d')}")

        # Loop through each work order
        cSync.loopOrders(work_orders)

        if last_week:
            break

        # Update the week start and week end dates for the next iteration
        week_start += timedelta(days=7)
        week_end = min(week_start + timedelta(days=7), stop_date)

    cSync.finalLogging()
