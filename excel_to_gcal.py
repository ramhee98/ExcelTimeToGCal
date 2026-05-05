import os
import sys

# Determine the directory of the script (portable, not hardcoded)
script_dir = os.path.dirname(os.path.realpath(__file__))

# Set working directory to script location
os.chdir(script_dir)

# Remove first sys.path entry if it's '', '/', or cwd
cwd = os.getcwd()
if sys.path[0] in ('', cwd, '/'):
    sys.path.pop(0)

# Insert the script directory explicitly (safe)
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

import os.path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import pandas as pd
from datetime import datetime, timedelta, time
from zoneinfo import ZoneInfo
import configparser

def load_config(path='config.ini'):
    config = configparser.ConfigParser()
    config.read(path)

    return {
        'calendar_id': config['calendar']['calendar_id'],
        'summary': config['calendar']['summary'],
        'description': config['calendar']['description'],
        'replace_event': config.getboolean('calendar', 'replace_event', fallback=False),
        'excel_file': config['excel']['excel_file'],
        'days_back': int(config['excel']['days_back'])
    }

def get_calendar_service(pwd, token_file="token.json", credentials_file="credentials.json"):
    """
    Handles Google Calendar API authentication and returns a service object.

    Returns:
        googleapiclient.discovery.Resource: Authenticated Google Calendar API service.
    """
    # If modifying these SCOPES, delete token.json
    SCOPES = ['https://www.googleapis.com/auth/calendar.events']
    token_file = pwd + "/" + token_file
    credentials_file = pwd + "/" + credentials_file

    creds = None
    # Load token if it exists
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)

    # If token is invalid or missing, log in and save new token
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as refresh_error:
                # Drop the now-useless token so the next run re-prompts via
                # OAuth, but don't crash if it was already removed and don't
                # swallow KeyboardInterrupt/SystemExit (the bare except did).
                try:
                    os.remove(token_file)
                except FileNotFoundError:
                    pass
                except OSError as remove_error:
                    print(f"Could not remove stale token file '{token_file}': {remove_error}")
                print(f"Token refresh failed ({refresh_error}); please renew.")
                creds = None
        if not creds:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the new token
        with open(token_file, 'w') as token_fp:
            token_fp.write(creds.to_json())

    # Build and return the calendar service
    service = build('calendar', 'v3', credentials=creds)
    return service

def create_event(service, calendar_id, summary, description, start_datetime, end_datetime, timezone='Europe/Zurich', replace_event=True):
    """
    Creates a new event in Google Calendar unless an event with the same description exists on the same day.
    """
    # The Excel-derived datetimes are naive local times. Attach the configured
    # timezone before sending them to the Calendar API so timeMin/timeMax are
    # interpreted in the user's zone instead of being silently treated as UTC.
    tz = ZoneInfo(timezone)
    start_of_day = start_datetime.replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=tz)
    end_of_day = start_of_day + timedelta(days=1)

    # Check for existing events on the same day
    events_result = service.events().list(
        calendarId=calendar_id,
        timeMin=start_of_day.isoformat(),
        timeMax=end_of_day.isoformat(),
        singleEvents=True,
        orderBy='startTime'
    ).execute()

    for event in events_result.get('items', []):
        if event.get('description', '') == description:
            if replace_event:
                service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()
                print(f"Deleted existing event: {event.get('summary')} on {start_datetime.date()}")
            else:
                print(f"Skipped existing event on {start_datetime.date()}: '{description}'")
                return None


    # Create event if no match found
    event = {
        'summary': summary,
        'description': description,
        'start': {
            'dateTime': start_datetime.isoformat(),
            'timeZone': timezone,
        },
        'end': {
            'dateTime': end_datetime.isoformat(),
            'timeZone': timezone,
        },
    }

    created_event = service.events().insert(calendarId=calendar_id, body=event).execute()
    print(f"Event created: {created_event.get('htmlLink')}")
    return created_event

def parse_workdays_from_dataframe(df, days_back=None):
    """
    Parses a DataFrame where each column is a day and each row is a labeled field (e.g., 'Datum', 'Start', etc.).
    Only includes past entries (no future dates). Optionally filters by days_back.
    
    Parameters:
        df (pandas.DataFrame): The input DataFrame with labeled rows.
        days_back (int or None): If set, only include dates within the last X days.

    Returns:
        List of dicts with 'description', 'start', and 'end' datetime objects.
    """
    entries = []
    today = datetime.now().date()
    date_cutoff = today - timedelta(days=days_back) if days_back else None

    # Verify the rows the parser expects to find actually exist; without this
    # a typo or renamed row would crash with a KeyError mid-loop.
    required_rows = ('Datum', 'Start', 'Ist zeit')
    missing = [row for row in required_rows if row not in df.index]
    if missing:
        print(f"Skipping sheet: missing required row(s): {', '.join(missing)}")
        return entries

    for col in df.columns:
        date_raw = df.loc['Datum', col]
        if pd.isna(date_raw) or str(date_raw).strip() == "":
            continue  # Skip if date is missing or empty
        date = pd.to_datetime(date_raw).date()
        if date > today:
            continue  # skip future dates
        if date_cutoff and date < date_cutoff:
            continue  # skip too-old entries

        ist_raw = df.loc['Ist zeit', col]
        if pd.isna(ist_raw) or str(ist_raw).strip() == "":
            continue  # No worked-hours value recorded for this day
        try:
            ist_zeit = float(ist_raw)
        except (TypeError, ValueError):
            print(f"Skipping column {col!r}: 'Ist zeit' is not numeric ({ist_raw!r})")
            continue
        if ist_zeit <= 0:
            continue  # Skip non-positive durations

        start_raw = df.loc['Start', col]
        if pd.isna(start_raw) or str(start_raw).strip() == "":
            continue

        # Only call pd.to_datetime() if needed
        if isinstance(start_raw, datetime):
            start_time = start_raw.time()
        elif isinstance(start_raw, str):
            start_time = pd.to_datetime(start_raw).time()
        elif isinstance(start_raw, pd.Timestamp):
            start_time = start_raw.to_pydatetime().time()
        elif isinstance(start_raw, time):
            start_time = start_raw
        else:
            raise TypeError(f"Unsupported start time format: {type(start_raw)}")

        start_datetime = datetime.combine(date, start_time)
        end_datetime = start_datetime + timedelta(hours=ist_zeit)

        entries.append({
            'start': start_datetime,
            'end': end_datetime,
        })

    return entries

def parse_all_sheets(xls_dict, days_back=None):
    """
    Takes a dict of DataFrames (from read_excel with sheet_name=None) and parses all sheets.
    """
    all_entries = []
    for sheet_name, df in xls_dict.items():
        entries = parse_workdays_from_dataframe(df, days_back=days_back)
        all_entries.extend(entries)
    return all_entries

def main():
    # get current path
    pwd = str(os.path.dirname(os.path.abspath(__file__)))

    #load configuration
    cfg = load_config(pwd + "/config.ini")
    calendar_id = cfg['calendar_id']
    summary = cfg['summary']
    description = cfg['description']
    replace_event = cfg['replace_event']
    excel_file = cfg['excel_file']
    days_back = cfg['days_back']

    # connect service
    service = get_calendar_service(pwd)

    # Verify the configured Excel file exists before talking to Google so a
    # bad path fails fast with a clear message instead of a pandas traceback.
    if not os.path.isfile(excel_file):
        print(f"❌ Excel file not found: {excel_file}")
        print("   Update 'excel_file' in config.ini to point at an existing .xlsx file.")
        return

    # reading events from excel
    xls = pd.read_excel(excel_file, sheet_name=None, index_col=0)
    entries = parse_all_sheets(xls, days_back)

    # create events
    for entry in entries:
        create_event(service, calendar_id, summary, description, entry['start'], entry['end'], "Europe/Zurich", replace_event)

main()