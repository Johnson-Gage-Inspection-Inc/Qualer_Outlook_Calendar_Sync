import logging
from datetime import datetime as dt
from datetime import time, timedelta
import os
import traceback
from tqdm import tqdm
import app.exceptions as ex
from app.outlook import Outlook
from app.qualer_api import QualerAPI

# Set the current directory to the directory of the script
new_directory = os.path.dirname(os.path.abspath(__file__))
os.chdir(new_directory)

# Set up logging
logging.basicConfig(filename='app/exception.log', level=logging.DEBUG)


class CalendarSync:
    def __init__(self):
        self.exceptions = []
        self.created_events = []
        self.deleted_events = []
        self.updated_events = []
        self.qualer = QualerAPI()
        self.outlook = Outlook()
        self.processor = CalendarSyncProcessor(self, self.qualer, self.outlook)
        self.getLastLog()

    def getLastLog(self):
        """Get the last log time from the log file"""
        if not (last_log := ex.get_last_log_time()):
            # TODO: Set up a way to notify myself when the last log time is None (email)
            raise SystemExit("Last log time is None. Please check the log file.")
        try:
            dateTime = dt.strptime(last_log, "%Y-%m-%d %H:%M:%S,%f")
            self.start_date = dateTime.replace(hour=0, minute=0, second=0, microsecond=0)
        except ValueError as e:
            logging.critical(f"Error parsing last log time: {e}")
            raise SystemExit("Error parsing last log time. Please check the log file.")

    def tickCounters(cSync, order):
        if order.status == "Past":
            return True
        if order.status == "Skipped":
            cSync.skipped_events.append(order.CustomOrderNumber)
            return True
        elif order.status == "Cancelled":
            cSync.deleted_events.append(order.CustomOrderNumber)
        elif order.status == "Updated":
            cSync.updated_events.append(order.CustomOrderNumber)
        elif order.status == "Created":
            cSync.created_events.append(order.CustomOrderNumber)

    def finalLogging(cSync):
        ex.group_orders_by_exception(cSync.exceptions)  # log the exceptions
        logging.info(f"Successfully created orders: {cSync.created_events}") if cSync.created_events else None
        logging.info(f"Successfully updated orders: {cSync.updated_events}") if cSync.updated_events else None
        logging.info(f"Successfully deleted orders: {cSync.deleted_events}") if cSync.deleted_events else None

        # Print the summary of CustomOrderNumbers count per unique exception
        for exception, count in ex.count_exceptions(cSync.exceptions).items():
            logging.info(f"Exception: {exception} (Count: {count})")


class CalendarSyncProcessor:
    def __init__(self, calendar_sync: CalendarSync, qualer_api: QualerAPI, outlook: Outlook):
        self.calendar_sync = calendar_sync
        self.qualer_api = qualer_api
        self.outlook = outlook

    def loopOrders(self):
        """Function to loop through work orders and process them"""
        self.qualer_api.get_work_orders(week_start, week_end)
        logging.info(f"Found {len(self.qualer_api.work_orders)} records between {week_start.strftime('%Y-%m-%d')} and {week_end.strftime('%Y-%m-%d')}")

        for work_order in tqdm(self.qualer_api.work_orders, desc="Processing Orders", leave=False, dynamic_ncols=True):
            try:
                CustomOrderNumber = work_order["CustomOrderNumber"][6:]
                order = QualerOrder(work_order)
                order.status = order.process_order(self.outlook)
                if self.calendar_sync.tickCounters(order):
                    continue

            except ValueError:
                self.calendar_sync.exceptions.append([CustomOrderNumber, traceback.format_exc()])
            except Exception as e:
                self.calendar_sync.exceptions.append([CustomOrderNumber, str(e)])


class DateTimeUtils:
    @staticmethod
    def combine_date_and_time(order: 'QualerOrder'):
        """Function to combine date and time for an order. If dates are missing, raises an exception. If times are both missing, assumes all day event."""
        required_fields = ["RequestFromTime", "RequestToTime", "RequestFromDate", "RequestToDate"]
        # If no values are missing, assume the event is not all day
        if all(isinstance(order.get(field), str) for field in required_fields):
            order.start_time = dt.combine(parse_datetime(order["RequestFromDate"]).date(), parse_datetime(order["RequestFromTime"]).time())
            order.end_time = dt.combine(parse_datetime(order["RequestToDate"]).date(), parse_datetime(order["RequestToTime"]).time())

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
            # Use default times if only one of the times is missing
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


class QualerOrder(dict):
    def __init__(order, *args, **kwargs):
        super(QualerOrder, order).__init__(*args, **kwargs)
        requiredKeys = ["ServiceOrderId", "CustomOrderNumber", "OrderStatus"]  # TODO: Add more required keys
        assert all(key in order for key in requiredKeys), f"Missing required keys: {', '.join([key for key in requiredKeys if key not in order])}"
        order.id = order["ServiceOrderId"]
        order.number = order["CustomOrderNumber"]
        order.status = order["OrderStatus"]
        order.qualer_api = QualerAPI()
        order.read_body()
        order.start_time = None
        order.end_time = None
        order.is_all_day = False
        order.combine_date_and_time()

    def process_order(order, outlook: Outlook, is_live=False):
        """Function to process an order"""
        order.CustomOrderNumber = int(order["CustomOrderNumber"][6:])
        event_id = outlook.check(order["ServiceOrderId"], order["CustomOrderNumber"])
        if order.get("RequestToDate") is None:
            return None

        request_to_date = parse_datetime(order["RequestToDate"]).date()
        if request_to_date < dt.now().date():
            return "Past"  # Skip the order if it has passed the RequestToDate

        if order["OrderStatus"] == "Cancelled":
            if event_id:
                outlook.event.delete(event_id) if is_live else logging.debug(f"Would have deleted event for {order['CustomOrderNumber']} if live")
                return "Cancelled"
            else:
                return "Skipped"  # Skip the cancelled order if it does not have an event in Outlook.

        if event_id:
            event_obj = outlook.event.lookup(event_id)
            if (differing_keys := diff(event_obj, dict(order))):  # NOTE: May be able to use `order` instead of `dict(order)`
                outlook.event.update(event_id, order, differing_keys == ['attendees']) if is_live else logging.debug(f"Would have updated event for {order['CustomOrderNumber']} if live")
                return "Updated"
            else:
                logging.info(f"Event for {order['CustomOrderNumber']} is up to date")
                return "Skipped"  # Skip if the changes to the order are irrelevant to the calendar event
        else:
            outlook.event.create(order) if is_live else logging.debug(f"Would have created event for {order['CustomOrderNumber']} if live")
            return "Created"

    def combine_date_and_time(order):
        DateTimeUtils.combine_date_and_time(order)

    def __dict__(order) -> dict:
        """Function to map a Qualer order to an Outlook event JSON"""
        return {
            "subject": order["ClientCompanyName"],
            "bodyPreview": order.number,
            "allowNewTimeProposals": False,
            "isAllDay": order.is_all_day,
            "categories": [],
            "showAs": "tentative" if order.status == "Scheduling" else "busy" if order.status == "Processing" else "free",
            "responseRequested": False,
            "isReminderOn": False,
            "isCancelled": True if order.status == "Cancelled" else False,
            "body": {
                "contentType": "html",
                "content": order.body(order.qualer_api.count_assets(order.id))
            },
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
        }

    @staticmethod
    def read_body():
        """Function to read the contents of body.html and return it as a string"""
        try:
            html_file_path = os.path.join(os.getcwd(), "app", "body.html")
            with open(html_file_path, 'r') as file:
                return file.read()
        except FileNotFoundError:
            logging.critical("File /app/body.html does not exist.")

    def body(order, asset_count):
        hyperlink = f'<a href="https://jgiquality.qualer.com/ServiceOrder/Info/{order.id}">{order.number}</a>'
        body_content = f"<b>{hyperlink}<br> Number of Assets:</b> {asset_count}"
        body = order.read_body().replace('<p class="MsoNormal"></p>', f'<p class="MsoNormal">{body_content}</p>')
        return body

    def extractAddressStr(order):
        """Function to extract the address string for an order"""
        address = order["ShippingAddress"]
        return f"{address['Address1']}, {address['City']}, {address['StateProvinceAbbreviation']} {address['ZipPostalCode']}"

    def gatherAssignees(order):
        assignees = []
        assignments = order.qualer_api.get_work_order_assignments(order.id)
        for assignment in assignments:
            if assignment["EmployeeId"] in assignees:
                continue
            try:
                assignees.append(order.qualer_api.prepare_event_attendee(assignment))
            except Exception as e:
                logging.error(e)
        return assignees


def parse_datetime(datetime_str):
    """Function that parses datetimes for combine_date_and_time()"""
    return dt.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S")


def diff(event1: dict, event2: dict):
    """Function to compare two event objects and return a list of keys whose values differ between the two events"""
    # Note: 'body' is not listed among these keys, because there are too many formatting discrepancies
    keys = ['subject', 'bodyPreview', 'allowNewTimeProposals', 'isAllDay', 'categories', 'showAs', 'responseRequested', 'isReminderOn', 'isCancelled', 'start', 'end', 'location', 'attendees']
    return [key for key in keys if event1[key] != event2.get(key)]


if __name__ == "__main__":
    cSync = CalendarSync()

    start_date = cSync.start_date
    stop_date = dt.now()  # Use the current date as the stop date

    week_start = start_date
    week_end = start_date + timedelta(days=7)

    total_weeks = (stop_date - start_date).days // 7 + 1
    pbar = tqdm(total=total_weeks, desc="Iterating weeks", unit="weeks", dynamic_ncols=True)

    while week_start <= stop_date:
        if week_end > stop_date:
            week_end = stop_date

        pbar.update(1)

        cSync.processor.loopOrders()

        week_start += timedelta(days=7)
        week_end = min(week_start + timedelta(days=7), stop_date)

    cSync.finalLogging()
    pbar.close()
