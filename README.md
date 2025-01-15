# Qualer Outlook Calendar Sync

This project synchronizes calendar events between Qualer and Outlook. It automates the process of keeping both calendars up-to-date with the latest events.

## Features

- Sync events from Qualer to Outlook
- Sync events from Outlook to Qualer
- Error logging for failed sync attempts

## Requirements

- Python 3.8+
- Microsoft Outlook
- Qualer account

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/Qualer_Outlook_Calendar_Sync.git
    cd Qualer_Outlook_Calendar_Sync
    ```

2. Create and activate a virtual environment:
    ```sh
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```

3. Install the required packages:
    ```sh
    pip install -r requirements.txt
    ```

4. Set up environment variables:
    - Create a `.env` file in the root directory with the following content:
        ```properties
        QUALER_USER = "your_qualer_email"
        QUALER_PASSWORD = "your_qualer_password"
        ```

## Usage

1. Run the synchronization script:
    ```sh
    python app/outlook.py
    ```

2. Check the `app/exception.log` file for any errors during the sync process.

## Logging

All exceptions and important logs are recorded in the `app/exception.log` file. This helps in debugging and understanding any issues that occur during the synchronization process.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## License

This project is licensed under the MIT License.