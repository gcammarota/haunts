import datetime
import json
import os
import sys
import string
import time

import gitlab
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

from .ini import get
from . import actions
from .calendars import create_event, delete_event, ORIGIN_TIME

# If modifying these scopes, delete the sheets-token file
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
GITLAB_TOKEN_CONFIG = "gitlab-token.json"
creds = None


def get_col(row, index):
    try:
        return row[index]
    except IndexError:
        return None


def get_credentials(config_dir):
    global creds
    if creds is not None:
        return
    # The token stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    token = config_dir / "sheets-token.json"
    credentials = config_dir / "credentials.json"
    if token.is_file():
        creds = Credentials.from_authorized_user_file(token.resolve(), SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials.resolve(), SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token.resolve(), "w") as token_file:
            token_file.write(creds.to_json())


def get_headers(sheet, month, indexes=False):
    """Scan headers of a month and returns a structure that assign headers names to indexes"""
    selected_month = (
        sheet.values()
        .get(spreadsheetId=get("CONTROLLER_SHEET_DOCUMENT_ID"), range=f"{month}!A1:ZZ1")
        .execute()
    )
    values = selected_month["values"][0]
    if indexes:
        return {k: values.index(k) for k in values}
    return {k: string.ascii_lowercase.upper()[values.index(k)] for k in values}


def sync_events(config_dir, sheet, data, calendars, projects, days, month):
    """Enumerate every data in the sheet.
    Create an event when action column is empty
    """
    headers = get_headers(sheet, month)
    headers_id = get_headers(sheet, month, indexes=True)
    last_to_time = None
    last_date = None

    for y, row in enumerate(data["values"]):
        current_date = get_col(row, headers_id["Date"])
        if not current_date:
            continue
        calendar = get_col(row, headers_id["Calendar"])
        if not calendar:
            continue
        date = ORIGIN_TIME + datetime.timedelta(days=current_date)

        # In case we changed day, let's restart from START_TIME
        if current_date != last_date:
            last_to_time = None
        last_date = current_date

        # short circuit for date filters
        skip = len(days) > 0
        for d in days:
            if date.date() == d.date():
                skip = False
                break
        if skip:
            continue

        calendar_id = None
        try:
            calendar_id = calendars[calendar]
        except KeyError:
            print(f"Cannot find a calendar id associated to calendar \"{calendar}\"")
            sys.exit(1)
        attendees = []
        attendees_cell_text = get_col(row, headers_id["Attendees"])
        if "Attendees" in headers_id and attendees_cell_text is not None:
            attendees = [
                {"email": attendee.strip()}
                for attendee in attendees_cell_text.split(",")
            ]

        try:
            action = row[headers_id["Action"]]
            if action == actions.IGNORE:
                continue
            elif action == actions.DELETE:
                delete_event(
                    config_dir=config_dir,
                    calendar=calendar_id,
                    event_id=get_col(row, headers_id["Event id"]),
                )
                print(f'Deleted event "{get_col(row, headers_id["Activity"])}"')
                request = sheet.values().batchClear(
                    spreadsheetId=get("CONTROLLER_SHEET_DOCUMENT_ID"),
                    body={
                        "ranges": [
                            f"{month}!{headers['Event id']}{y + 2}",
                            f"{month}!{headers['Link']}{y + 2}",
                            f"{month}!{headers['Action']}{y + 2}",
                        ],
                    },
                )

                try:
                    request.execute()
                except HttpError as err:
                    if err.status_code == 429:
                        print("Too many requests")
                        print(err.error_details)
                        print("haunts will now pause for a while ⏲…")
                        time.sleep(60)
                        print("Retrying…")
                        request.execute()
                    else:
                        raise

                continue
            else:
                # There's something in the action cell, but not recognized
                print(f"Unknown action {action}. Ignoring…")
                continue
        except IndexError:
            # We have no data there
            pass

        event = create_event(
            config_dir=config_dir,
            calendar=calendar_id,
            date=date,
            summary=get_col(row, headers_id["Activity"]),
            details=get_col(row, headers_id["Details"]),
            start_time=get_col(row, headers_id["Start"]),
            stop_time=get_col(row, headers_id["Stop"]),
            attendees=attendees,
            from_time=last_to_time,
        )
        last_to_time = event["next_slot"]

        # Save the event id, required to interact with the event in future
        request = sheet.values().update(
            spreadsheetId=get("CONTROLLER_SHEET_DOCUMENT_ID"),
            range=f"{month}!{headers['Action']}{y + 2}",
            valueInputOption="RAW",
            body={"values": [[actions.IGNORE]]},
        )
        request.execute()

        # Save the event id, required to interact with the event in future
        request = sheet.values().update(
            spreadsheetId=get("CONTROLLER_SHEET_DOCUMENT_ID"),
            range=f"{month}!{headers['Event id']}{y + 2}",
            valueInputOption="RAW",
            body={"values": [[event["id"]]]},
        )
        request.execute()

        # Quick link to the event on the calendar
        request = sheet.values().update(
            spreadsheetId=get("CONTROLLER_SHEET_DOCUMENT_ID"),
            range=f"{month}!{headers['Link']}{y + 2}",
            valueInputOption="USER_ENTERED",
            body={"values": [[f"=HYPERLINK(\"{event['link']}\";\"open\")"]]},
        )
        request.execute()

        project = get_col(row, headers_id["Project"])
        issue = get_col(row, headers_id["Issue"])
        try:
            pid = projects[project]
        except KeyError:
            print(f"Cannot find a project id, skipping comment to issue '{issue}'.")
            continue
        spent = get_col(row, headers_id["Spent"])
        if not spent:
            spent = 8
        try:
            url, gitlab_token = read_gitlab_token(config_dir)
        except ValueError as e:
            print(f"ValueError: {e}")
            continue
        add_spent_time_on_gitlab_issue(
            url,
            gitlab_token,
            pid,
            issue.strip("#"),
            spent
        )
        print(f"Added {spent} hours to issue {project}{issue}")


def read_gitlab_token(config_dir):
    if GITLAB_TOKEN_CONFIG not in os.listdir(config_dir):
        raise ValueError("No GitLab token configuration")
    token_file = config_dir / GITLAB_TOKEN_CONFIG
    with open(token_file) as f:
        info = json.load(f)
    return info["url"], info["token"]


def add_spent_time_on_gitlab_issue(gitlab_base_url, private_token, project_id, issue_id, spent):
    # Authenticate to the GitLab API
    gl = gitlab.Gitlab(gitlab_base_url, private_token=private_token)

    # Get the issue object
    project = gl.projects.get(project_id)
    issue = project.issues.get(issue_id)

    # Add a comment to the issue
    issue.notes.create({'body': f"/spend {spent}h"})


def get_calendars(sheet):
    RANGE = f"{get('CONTROLLER_SHEET_NAME', 'config')}!A2:B"
    calendars = (
        sheet.values()
        .get(spreadsheetId=get("CONTROLLER_SHEET_DOCUMENT_ID"), range=RANGE)
        .execute()
    )
    values = calendars.get("values", [])
    return {alias: id for [id, alias] in values}


def get_projects(sheet):
    RANGE = f"{get('CONTROLLER_SHEET_NAME', 'projects')}!A2:B"
    projects = (
        sheet.values()
        .get(spreadsheetId=get("CONTROLLER_SHEET_DOCUMENT_ID"), range=RANGE)
        .execute()
    )
    values = projects.get("values", [])
    return {alias: id for [alias, id] in values}


def sync_report(config_dir, month, days=[]):
    """Open a sheet, analyze it and populate calendars with new events"""
    # The ID and range of the controller timesheet
    get_credentials(config_dir)

    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()

    try:
        document_id = get("CONTROLLER_SHEET_DOCUMENT_ID")
    except KeyError:
        print(
            "A value for CONTROLLER_SHEET_DOCUMENT_ID is required but "
            "is not specified in your ini file"
        )
        sys.exit(1)

    if month is None:
        sheets = sheet.get(spreadsheetId=document_id).execute()
        month = sheets["sheets"][-1]["properties"]["title"]
    print("Sheet: {}".format(month))
    data = (
        sheet.values()
        .get(
            spreadsheetId=document_id,
            range=f"{month}!A2:ZZ",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
    )

    calendars = get_calendars(sheet)
    projects = get_projects(sheet)
    sync_events(config_dir, sheet, data, calendars, projects, days=days, month=month)
