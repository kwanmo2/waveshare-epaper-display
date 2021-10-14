import datetime
import time
import pickle
import os.path
import os
import logging
from exchangelib import Credentials, Account, CalendarItem, Mailbox, Message
from exchangelib import items
from exchangelib.items import item
from datetime import timedelta

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import outlook_util
from utility import is_stale, update_svg, configure_logging

configure_logging()

# note: increasing this will require updates to the SVG template to accommodate more events
max_event_results = 4
credentials = Credentials('Exchange ID', 'Exchange PW')
account = Account('Exchange ID', credentials=credentials, autodiscover=True)

google_calendar_id = os.getenv("GOOGLE_CALENDAR_ID", "primary")
outlook_calendar_id = os.getenv("OUTLOOK_CALENDAR_ID", None)

ttl = float(os.getenv("CALENDAR_TTL", 1 * 60 * 60))
MaxCount = 5

def get_outlook_events(max_event_results):

    outlook_calendar_pickle = 'outlookcalendar.pickle'

    if is_stale(os.getcwd() + "/" + outlook_calendar_pickle, ttl):
        logging.debug("Pickle is stale, calling the Outlook Calendar API")
        now_iso = datetime.datetime.now().astimezone().replace(microsecond=0).isoformat()
        oneyearlater_iso = (datetime.datetime.now().astimezone()
                            + datetime.timedelta(days=365)).astimezone().isoformat()
        
        calendar_item =[]

        
        
        access_token = outlook_util.get_access_token()
        events_data = outlook_util.get_outlook_calendar_events(
                                                                outlook_calendar_id,
                                                                now_iso,
                                                                oneyearlater_iso,
                                                                access_token)
        logging.debug(events_data)

        with open(outlook_calendar_pickle, 'wb') as cal:
            pickle.dump(events_data, cal)
    else:
        logging.info("Found in cache")
        with open(outlook_calendar_pickle, 'rb') as cal:
            events_data = pickle.load(cal)

    return events_data


def get_output_dict_from_outlook_events(outlook_events, event_slot_count):
    events = outlook_events["value"]
    formatted_events = {}
    event_count = len(events)
    for event_i in range(event_slot_count):
        event_label_id = str(event_i + 1)
        if (event_i <= event_count - 1):
            formatted_events['CAL_DATETIME_' + event_label_id] = outlook_util.get_outlook_datetime_formatted(events[event_i])
            formatted_events['CAL_DESC_' + event_label_id] = events[event_i]['subject']
        else:
            formatted_events['CAL_DATETIME_' + event_label_id] = ""
            formatted_events['CAL_DESC_' + event_label_id] = ""
    return formatted_events


def get_google_credentials():

    google_token_pickle = 'token.pickle'
    google_credentials_json = 'credentials.json'
    google_api_scopes = ['https://www.googleapis.com/auth/calendar.readonly']

    credentials = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(google_token_pickle):
        with open(google_token_pickle, 'rb') as token:
            credentials = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                google_credentials_json, google_api_scopes)
            credentials = flow.run_local_server()
        # Save the credentials for the next run
        with open(google_token_pickle, 'wb') as token:
            pickle.dump(credentials, token)

    return credentials


def get_google_events(max_event_results):

    google_calendar_pickle = 'calendar.pickle'

    service = build('calendar', 'v3', credentials=get_google_credentials(), cache_discovery=False)

    events_result = None

    if is_stale(os.getcwd() + "/" + google_calendar_pickle, ttl):
        logging.debug("Pickle is stale, calling the Calendar API")

        # Call the Calendar API
        events_result = service.events().list(
            calendarId=google_calendar_id,
            timeMin=datetime.datetime.utcnow().isoformat() + 'Z',
            maxResults=max_event_results,
            singleEvents=True,
            orderBy='startTime').execute()

        with open(google_calendar_pickle, 'wb') as cal:
            pickle.dump(events_result, cal)

    else:
        logging.info("Found in cache")
        with open(google_calendar_pickle, 'rb') as cal:
            events_result = pickle.load(cal)

    events = events_result.get('items', [])

    if not events:
        logging.info("No upcoming events found.")

    return events


def get_output_dict_from_google_events(events, event_slot_count):
    formatted_events = {}
    event_count = len(events)
    for event_i in range(event_slot_count):
        event_label_id = str(event_i + 1)
        if (event_i <= event_count - 1):
            formatted_events['CAL_DATETIME_' + event_label_id] = get_google_datetime_formatted(events[event_i]['start'])
            formatted_events['CAL_DESC_' + event_label_id] = events[event_i]['summary']
        else:
            formatted_events['CAL_DATETIME_' + event_label_id] = ""
            formatted_events['CAL_DESC_' + event_label_id] = ""
    return formatted_events


def get_google_datetime_formatted(event_start):
    if(event_start.get('dateTime')):
        start = event_start.get('dateTime')
        day = time.strftime("%a %b %-d, %-I:%M %p", time.strptime(start, "%Y-%m-%dT%H:%M:%S%z"))
    else:
        start = event_start.get('date')
        day = time.strftime("%a %b %-d", time.strptime(start, "%Y-%m-%d"))
    return day


#def get_exchange_events(max_event_results):

def get_output_dict_from_exchange_events(MaxCount):
    
    ##Get email from exchange account
    ttime = datetime.datetime.now()
    formatted_events= {}
    i = 0
    for item in account.inbox.filter(is_read=False).order_by('-datetime_received')[:5]:
      event_label_id = str(i+1)
      formatted_events['MAIL_'+event_label_id] = str(item.subject)
      i=i+1
    
    ##Get calendar from exchange account
    i = 0
    items = account.calendar.view(start=datetime.datetime(ttime.year, ttime.month, ttime.day, tzinfo=account.default_timezone),
        end=datetime.datetime(ttime.year, ttime.month, ttime.day, tzinfo=account.default_timezone) + datetime.timedelta(days=7))
    for item in items:
      event_label_id = str(i+1)
      formatted_events['CAL_DATETIME_' + event_label_id] = str(item.start+timedelta(hours=9)).split('+')[0][0:16]
      formatted_events['CAL_DESC_' + event_label_id] = str(item.subject)
      i=i+1
    i=0
    
    return formatted_events


def main():

    output_svg_filename = 'screen-output-weather.svg'
    """
    if outlook_calendar_id:
        logging.info("Fetching Outlook Calendar Events")
        outlook_events = get_outlook_events(max_event_results)
        output_dict = get_output_dict_from_outlook_events(outlook_events, max_event_results)

    else:
        logging.info("Fetching Google Calendar Events")
        google_events = get_google_events(max_event_results)
        output_dict = get_output_dict_from_google_events(google_events, max_event_results)
    """
    output_dict = get_output_dict_from_exchange_events(MaxCount)
    logging.info("Fetching Exchange Calendar Events")
    
    logging.info("main() - {}".format(output_dict))

    logging.info("Updating SVG")
    update_svg(output_svg_filename, output_svg_filename, output_dict)


if __name__ == "__main__":
    main()
