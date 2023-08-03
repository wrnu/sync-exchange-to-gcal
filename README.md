# SyncExToGCal

## Description

SyncExToGCal is a Python script designed to synchronize events from a Microsoft Exchange Calendar with a Google Calendar. The process includes setting up necessary configurations and establishing connections with both Exchange and Google Calendar servers.

In the synchronization phase, the script fetches events within a specified time range from the Exchange Calendar, and either creates these events on the Google Calendar or updates them if they already exist. If an event already exists on the Google Calendar, the script checks for any changes on the Exchange Calendar side and updates the Google Calendar event accordingly.

Importantly, the script also handles event deletions, but only for the events it has created. This means if an event, originally created by this script, is removed from the Exchange Calendar, it is also deleted from the Google Calendar.

Furthermore, events can be selectively ignored during synchronization based on their titles, offering an added layer of control. All operations, whether creation, updating, or deletion, are conducted while preserving the time zone of the original event from the Exchange Calendar.

In essence, SyncExToGCal provides a robust, selective synchronization solution, ensuring your Google Calendar stays up-to-date with your Exchange Calendar in a controlled manner.

## Setup

1. **Clone the repository.**

```bash

git clone https://github.com/wrnu/sync-exchange-to-gcal.git
cd sync-exchange-to-gcal
```

1. **Set up a virtual environment (optional, but recommended).**

```bash

python3 -m venv venv
source venv/bin/activate  # On Windows use `.\venv\Scripts\activate`
```

1. **Install the requirements.**

```bash

pip install -r requirements.txt
```

1. **Set up your environment variables.**

Create a `.env` file in the root directory and populate it with the necessary environment variables. Refer to the [Environment Variables](https://chat.openai.com/?model=gpt-4#environment-variables)  section for more details.

## Usage

To run the script, use the following command:

```bash

python main.py
```

The script will then synchronize events from your Exchange Calendar to your Google Calendar according to the configuration parameters you've set.

## Environment Variables

The script requires the following environment variables:

- `EX2GCAL_NUM_DAYS_TO_SYNC` - The number of days to synchronize (default is "1").
- `EX2GCAL_EVENT_TITLE_PREFIX` - A prefix to add to event titles (default is "").
- `EX2GCAL_EVENT_TITLES_TO_SKIP` - Event titles to skip during synchronization. Should be comma-separated string (default is "").
- `EWS_EMAIL_ADDRESS` - Your Exchange email address.
- `EWS_PASSWORD` - Your Exchange password.
- `EWS_SERVER` - Your Exchange server.

You also need to set up Google API credentials and save them as `credentials.json` in the project's root directory. Visit [this page](https://developers.google.com/calendar/quickstart/python)  to set up the credentials.

## Note

Please note that this script is designed to work with basic Exchange Calendar to Google Calendar synchronization. If you have complex calendar events or specific requirements, you may need to modify the script accordingly. The code includes detailed comments to assist you with this process.
