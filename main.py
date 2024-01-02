import os
import pickle
from datetime import timedelta
import logging
import re

from bs4 import BeautifulSoup
from dateutil.parser import parse
from dotenv import load_dotenv
from exchangelib import (DELEGATE, Account, Configuration, Credentials,
                         EWSDateTime, EWSTimeZone)
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Load environment variables from .env file
load_dotenv()

# Set up a logger for your module
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# Create a handler for the logger to specify where the logs should go (e.g., console, file)
handler = logging.StreamHandler()  # Sends logs to the console
handler.setLevel(logging.INFO)

# Create a formatter to define the log message format
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)

# Add the handler to the logger
logger.addHandler(handler)


def sanitize_for_google_calendar(html):
    """
    Sanitizes HTML for Google Calendar.
    """
    allowed_tags = ['b', 'i', 'u', 'a', 'br', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
                    'img', 'blockquote', 'ol', 'ul', 'li', 'em', 'strong', 'code', 'hr']
    allowed_attributes = {
        'a': ['href'],
        'img': ['src', 'alt']
    }

    # Remove comments using regular expressions
    html = re.sub('<!--.*?-->', '', html, flags=re.DOTALL)

    # Further sanitization using BeautifulSoup
    soup = BeautifulSoup(html, 'lxml')  # Using lxml parser for better handling

    # Remove disallowed tags and attributes
    for tag in soup.find_all(True):
        if tag.name not in allowed_tags:
            tag.unwrap()
        else:
            tag.attrs = {name: value for name, value in tag.attrs.items()
                         if name in allowed_attributes.get(tag.name, [])}

    sanitized_html = str(soup).strip()

    return sanitized_html


class SyncExToGCal:
    """
    This class synchronizes events from an Exchange calendar to a Google Calendar.
    """

    def __init__(self):
        """
        Initializes the SyncExToGCal class.
        """
        self.setup_config()
        self.setup_exchange()
        self.setup_time()
        self.setup_google_calendar()

    def setup_config(self):
        """
        Sets up configuration parameters for the synchronization process.
        """
        self.num_days_to_sync = int(os.getenv("EX2GCAL_NUM_DAYS_TO_SYNC", "1"))
        self.event_title_prefix = os.getenv("EX2GCAL_EVENT_TITLE_PREFIX", "")
        self.event_titles_to_skip = os.environ.get(
            "EX2GCAL_EVENT_TITLES_TO_SKIP", "").split(',')
        logger.info('Event title prefix: ' + self.event_title_prefix)
        logger.info('Event titles to skip: ' +
                    ', '.join(self.event_titles_to_skip))

    def setup_exchange(self):
        """
        Sets up the Exchange server connection.
        """
        email_address = os.getenv("EWS_EMAIL_ADDRESS")
        password = os.getenv("EWS_PASSWORD")
        server = os.getenv("EWS_SERVER")

        try:
            credentials = Credentials(email_address, password)
            config = Configuration(server=server, credentials=credentials)
            self.exchange_account = Account(
                email_address, config=config, autodiscover=False, access_type=DELEGATE)
        except Exception as e:
            logger.error(f"Failed to setup exchange: {e}")

    def fetch_exchange_events(self):
        """
        Fetches events from the Exchange calendar within the specified time range.
        """
        try:
            self.exchange_events = self.exchange_account.calendar.view(
                start=self.start,
                end=self.end,
            )
        except Exception as e:
            logger.error(f"Failed to fetch Exchange events: {e}")

    def setup_time(self):
        """
        Sets up the time range for synchronization.
        """
        try:
            tz = EWSTimeZone.localzone()
            self.timezone = EWSTimeZone.MS_TO_IANA_MAP[tz.ms_id]
            now = EWSDateTime.now(tz)
            start = EWSDateTime(now.year, now.month, now.day, tzinfo=tz)
            end = start + timedelta(days=self.num_days_to_sync)
            self.start = start
            self.end = EWSDateTime(end.year, end.month,
                                   end.day, 23, 59, 59, tzinfo=tz)
            logger.info(f'Timezone: {tz.ms_id} -> {self.timezone}')
            logger.info(f'Sync start date: {self.start}')
            logger.info(f'Sync end date: {self.end}')
        except Exception as e:
            logger.error(f"Failed to setup time: {e}")

    def setup_google_calendar(self):
        """
        Sets up the Google Calendar API connection.
        """
        try:
            SCOPES = ['https://www.googleapis.com/auth/calendar']
            creds = None
            if os.path.exists('token.pickle'):
                with open('token.pickle', 'rb') as token:
                    creds = pickle.load(token)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'credentials.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                with open('token.pickle', 'wb') as token:
                    pickle.dump(creds, token)

            self.gcal_service = build('calendar', 'v3', credentials=creds)
        except Exception as e:
            logger.error(f"Failed to setup Google Calendar: {e}")

    def fetch_google_events(self):
        """
        Fetches events from the Google Calendar within the specified time range.
        """
        try:
            start_str = self.start.strftime("%Y-%m-%dT%H:%M:%S%z")
            end_str = self.end.strftime("%Y-%m-%dT%H:%M:%S%z")

            self.google_events = {}

            page_token = None
            while True:
                google_events_result = self.gcal_service.events().list(
                    calendarId='primary',
                    timeMin=start_str,
                    timeMax=end_str,
                    singleEvents=True,
                    orderBy='startTime',
                    pageToken=page_token
                ).execute()

                # Find and store all Google events that have been created by this script.
                # This is done by checking the extendedProperties.private.exchangeId property
                # which contains the Exchange event ID of the corresponding exchange event
                # Only events that have this property are managed by this script
                for g_event in google_events_result.get('items', []):
                    if g_event.get('extendedProperties', {}).get('private', {}).get('exchangeId'):
                        self.google_events[g_event.get('extendedProperties', {}).get(
                            'private', {}).get('exchangeId')] = g_event

                page_token = google_events_result.get('nextPageToken')
                if not page_token:
                    break
        except Exception as e:
            logger.error(f"Failed to fetch Google events: {e}")

    def create_or_update_google_event_from_exchange(self, item):
        """
        Manages and individual event by creating or updating it in the Google Calendar.
        """
        try:
            event_id = item.id
            event = {
                'summary': self.event_title_prefix + (item.subject[:1024] if item.subject else ''),
                'description': sanitize_for_google_calendar(item.body[:8192] if item.body else ''),
                'start': {
                    'dateTime': item.start.strftime("%Y-%m-%dT%H:%M:%S%z"),
                    'timeZone': self.timezone,
                },
                'end': {
                    'dateTime': item.end.strftime("%Y-%m-%dT%H:%M:%S%z"),
                    'timeZone': self.timezone,
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                        {'method': 'popup', 'minutes': 10},
                    ],
                },
                'extendedProperties': {
                    'private': {
                        'exchangeId': event_id
                    }
                },
                'transparency': 'opaque'  # Sets the event as "Busy"
            }

            if event_id in self.google_events:
                g_event = self.google_events[event_id]
                # Check if the Exchange event has changed
                if (
                    event['summary'] != g_event.get('summary') or
                    event['description'] != g_event.get('description') or
                    parse(event['start']['dateTime']) != parse(g_event['start'].get('dateTime')) or
                    parse(event['end']['dateTime']) != parse(
                        g_event['end'].get('dateTime'))
                ):
                    try:
                        self.gcal_service.events().update(calendarId='primary',
                                                          eventId=g_event['id'], body=event).execute()
                        logger.info(
                            f'Event updated: ({g_event["summary"]}) {g_event["htmlLink"]}')
                    except HttpError as error:
                        logger.error(f'Error: {error}, Event: {g_event}')

                del self.google_events[event_id]

            else:
                try:
                    g_event = self.gcal_service.events().insert(
                        calendarId='primary', body=event).execute()
                    logger.info(
                        f'Event created: ({g_event["summary"]}) {g_event["htmlLink"]}')
                except HttpError as error:
                    logger.error(f'Error: {error}, Event: {g_event}')

        except Exception as e:
            logger.error(f"Failed to manage events: {e}")

    def delete_google_events(self):
        """
        Deletes events from Google Calendar that are not present in the Exchange calendar but were created by this script.
        """
        try:
            for g_event in self.google_events.values():
                try:
                    self.gcal_service.events().delete(calendarId='primary',
                                                      eventId=g_event['id']).execute()
                    logger.info(
                        f'Event deleted: ({g_event["summary"]}) {g_event["htmlLink"]}')
                except HttpError as error:
                    logger.error(f'Error: {error}, Event: {g_event}')
        except Exception as e:
            logger.error(f"Failed to delete missing events: {e}")

    def sync(self):
        """
        Initiates the synchronization process.
        """
        try:
            self.fetch_google_events()
            self.fetch_exchange_events()

            # Create or update Google events from Exchange events
            for item in self.exchange_events:
                # Skip events with titles that are in the skip list
                if item.subject in self.event_titles_to_skip:
                    logger.info(f"Event skipped: {item.subject}")
                    continue

                self.create_or_update_google_event_from_exchange(item)

            # Delete Google events that were created by this script but are not present in the Exchange calendar
            self.delete_google_events()
        except Exception as e:
            logger.error(f"Failed to sync events: {e}")


if __name__ == '__main__':
    sync = SyncExToGCal()
    sync.sync()
