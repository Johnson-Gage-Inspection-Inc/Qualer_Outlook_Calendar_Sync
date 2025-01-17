# Qualer Outlook Calendar Sync

This project synchronizes calendar events between Qualer and Outlook. It ensures that events created or updated in Qualer are reflected in Outlook.

## Features

- Sync calendar events from Qualer to Outlook
- Handle event updates and deletions

## Requirements

- Python 3.8+
- pip

## Dependencies

- `requests (Apache-2.0)`: For making HTTP requests to the Qualer API.
- `bs4 (MIT License)`: For parsing and handling HTML responses.

## Setup

1. Clone the repository:
    ```sh
    git clone https://github.com/Johnson-Gage-Inspection-Inc/Qualer_Outlook_Calendar_Sync.git
    cd Qualer_Outlook_Calendar_Sync
    ```

2. Install the required dependencies:
    ```sh
    pip install -r requirements.txt
    ```

3. Configure environment variables:
    - Create a `.env` file in the root directory with the following content:
        ```properties
        QUALER_USER = your_qualer_email
        QUALER_PASSWORD = your_qualer_password
        ```

## Usage

1. Run the synchronization script:
    ```sh
    python calendar_sync.py
    ```

## Project Structure

- `calendar_sync.py`: Main script to synchronize calendars.
- `app/qualer_api.py`: Handles API interactions with Qualer.
- `app/outlook.py`: Handles API interactions with Outlook.
- `requirements.txt`: Lists the Python dependencies.
- `.env`: Contains environment variables for Qualer credentials.
- `.gitignore`: Specifies files and directories to be ignored by git.

## Logging

Logs are stored in `app/exception.log` to track synchronization issues and exceptions.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request.

## License

This project is licensed under the MIT License.
