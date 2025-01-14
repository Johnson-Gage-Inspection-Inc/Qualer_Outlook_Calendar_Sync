import logging

###########################################################################################
################################# Exception Handling ######################################
###########################################################################################


# Configure logging

log_file = 'app/exception.log'
logging.basicConfig(filename=log_file, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')


# Function to count the number of CustomOrderNumbers per unique exception
def count_exceptions(exceptions: list) -> dict:
    # Create a dictionary to store the count of CustomOrderNumbers per unique exception
    order_exceptions_count = {}

    # Iterate over the exceptions array
    for exception in exceptions:
        # Get the exception type
        exception_type = exception[1]

        # Check if the exception type already exists in the dictionary
        if exception_type in order_exceptions_count:
            order_exceptions_count[exception_type] += 1
        else:
            order_exceptions_count[exception_type] = 1
    return order_exceptions_count


# Function to group CustomOrderNumbers by exception type and log the output
def group_orders_by_exception(exceptions: list) -> None:
    order_exceptions = {}
    for order_number, exception in exceptions:
        if exception in order_exceptions:
            order_exceptions[exception].append(order_number)
        else:
            order_exceptions[exception] = [order_number]

    for exception_type, order_numbers in order_exceptions.items():
        logging.exception(f"Exception type: {exception_type}, Order numbers: {order_numbers}")


def get_last_log_time() -> str:
    """Extract the timestamp from the last log entry"""
    with open(log_file, 'r') as file:
        if lines := file.readlines():
            last_log_entry = lines[-1].strip()
            return last_log_entry.split(' - ')[0]
        else:
            raise logging.critical("Log file is empty")
