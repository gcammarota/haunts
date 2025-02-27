import datetime
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

from . import LOGGER
from .ini import get
from .credentials import get_credentials

LOCAL_TIMEZONE = datetime.datetime.utcnow().astimezone().strftime("%z")
# Weird google spreadsheet date management
ORIGIN_TIME = datetime.datetime.strptime(
    f"1899-12-30T00:00:00{LOCAL_TIMEZONE}", "%Y-%m-%dT%H:%M:%S%z"
)
# If modifying these scopes, delete the calendars-token file.
SCOPES = ["https://www.googleapis.com/auth/calendar"]


def init(config_dir):
    get_credentials(config_dir, SCOPES, "calendars-token.json")


def create_event(config_dir, calendar, date, summary, details, start_time, stop_time, attendees, from_time=None):
    creds = get_credentials(config_dir, SCOPES, "calendars-token.json")
    service = build("calendar", "v3", credentials=creds)

    from_time = from_time or get("START_TIME")
    today = datetime.datetime.strptime(
        f"{date.strftime('%Y-%m-%d')}T00:00:00{LOCAL_TIMEZONE}",
        "%Y-%m-%dT%H:%M:%S%z",
    )

    start = end = today
    today_str = today.strftime("%d/%m/%Y")
    spent_comment = "all-day"
    startParams = {"timeZone": LOCAL_TIMEZONE}
    endParams = {"timeZone": LOCAL_TIMEZONE}

    if start_time and stop_time:
        start = today + datetime.timedelta(days=start_time)
        startParams.update({"dateTime": start.isoformat()})
        end = today + datetime.timedelta(days=stop_time)
        endParams.update({"dateTime": end.isoformat()})
        spent_comment = f'{start.strftime("%H:%M")}' + " - " + f'{end.strftime("%H:%M")}'
    else:
        startParams.update({"date": start.isoformat()[:10]})
        endParams.update({"date": end.isoformat()[:10]})

    event = {
        "summary": summary,
        "description": details,
        "start": startParams,
        "end": endParams,
        "attendees": attendees,
    }

    LOGGER.debug(calendar, date, summary, details, start, end, event)
    event = service.events().insert(calendarId=calendar, body=event).execute()
    LOGGER.debug(event.items())
    print(
        f'Created event "{summary}" ({f"{today_str} {spent_comment}"}) on calendar {event["organizer"]["displayName"]}'
    )
    event_data = {
        "id": event["id"],
        "next_slot": end.strftime("%H:%M"),
        "link": event["htmlLink"],
    }
    return event_data


def execute(config_dir):
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = get_credentials(config_dir, SCOPES, "calendars-token.json")

    service = build("calendar", "v3", credentials=creds)

    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time
    print("Getting the upcoming 10 events")
    events_result = (
        service.events()
        .list(
            calendarId="c_bu7esjsc8qt8vc6gtjruuq94js@group.calendar.google.com",
            timeMin=now,
            maxResults=10,
            singleEvents=True,
            orderBy="startTime",
        )
        .execute()
    )
    events = events_result.get("items", [])

    if not events:
        print("No upcoming events found.")
    for event in events:
        start = event["start"].get("dateTime", event["start"].get("date"))
        print(start, event["summary"])


def delete_event(config_dir, calendar, event_id):
    creds = get_credentials(config_dir, SCOPES, "calendars-token.json")
    service = build("calendar", "v3", credentials=creds)
    if not event_id:
        print("Missing id. Skipping…")
        return
    try:
        service.events().delete(calendarId=calendar, eventId=event_id).execute()
    except HttpError as err:
        if err.status_code == 410:
            print("Event {event_id} already deleted")
