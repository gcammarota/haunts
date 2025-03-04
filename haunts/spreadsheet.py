import datetime
import json
import os
import string
import sys
import time

import gitlab
import rich
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from . import actions
from .calendars import ORIGIN_TIME, create_event, delete_event
from .credentials import get_credentials
from .ini import get

# If modifying these scopes, delete the sheets-token file
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
GITLAB_TOKEN_CONFIG = "gitlab-token.json"


def get_col(row, index):
    try:
        return row[index]
    except IndexError:
        return None


def get_first_empty_line(sheet, month):
    """Get the first empty line in a month."""
    RANGE = f"{month}!A1:A"
    lines = (
        sheet.values()
        .get(spreadsheetId=get("CONTROLLER_SHEET_DOCUMENT_ID"), range=RANGE)
        .execute()
    )
    values = lines.get("values", [])
    return len(values) + 1


def format_duration(duration):
    """Given a timedelta duration, format is as a string.

    String format will be H,X or H if minutes are 0.
    X is the decimal part of the hour (30 minutes are 0.5 hours, etc)
    """
    hours = duration.total_seconds() / 3600
    if hours % 1 == 0:
        return str(int(hours))
    return str(hours).replace(".", ",")


def try_request(request):
    try:
        request.execute()
    except HttpError as err:
        if err.status_code == 429:
            rich.print("Too many requests")
            rich.print(err.error_details)
            rich.print("haunts will now pause for a while ⏲…")
            time.sleep(60)
            rich.print("Retrying…")
            request.execute()
        else:
            raise


def append_line(
    sheet,
    month,
    date_col,
    start_col,
    stop_col,
    calendar_col,
    activity_col,
    event_id_col=None,
    link_col="",
    details_col="",
    action_col="",
    duration_col=None,
):
    """Append a new line at the end of a sheet."""
    next_av_line = get_first_empty_line(sheet, month)
    headers_id = get_headers(sheet, month, indexes=True)
    # Now write a new line at position next_av_line
    RANGE = f"{month}!A{next_av_line}:ZZ{next_av_line}"
    values_line = []
    formatted_time_col = start_col.strftime("%H.%M") if start_col else ""
    formatted_stop_col = stop_col.strftime("%H.%M") if stop_col else ""
    formatted_duration_col = format_duration(duration_col) if duration_col else ""
    full_day = formatted_time_col == "00:00" and formatted_duration_col == "24"
    for key, index in headers_id.items():
        if key == "Date":
            values_line.append(date_col.strftime("%d/%m/%Y"))
        elif key == get("START_TIME_COLUMN_NAME", "Start time"):
            values_line.append(formatted_time_col if not full_day else "")
        elif key == get("STOP_TIME_COLUMN_NAME", "Stop time"):
            values_line.append(formatted_stop_col if not full_day else "")
        elif key == get("CALENDAR_COLUMN_NAME", "Calendar"):
            values_line.append(calendar_col)
        elif key == get("ACTIVITY_COLUMN_NAME", "Title"):
            values_line.append(activity_col)
        elif key == get("DETAILS_COLUMN_NAME", "Details"):
            values_line.append(details_col)
        elif key == get("EVENT_ID_COLUMN_NAME", "Event id"):
            values_line.append(event_id_col)
        elif key == get("LINK_COLUMN_NAME", "Link"):
            values_line.append(link_col)
        elif key == get("ACTION_COLUMN_NAME", "Action"):
            values_line.append(action_col)
        elif key == get("SPENT_COLUMN_NAME", "Spent"):
            values_line.append(formatted_duration_col if not full_day else "")
        else:
            values_line.append("")

    request = sheet.values().batchUpdate(
        spreadsheetId=get("CONTROLLER_SHEET_DOCUMENT_ID"),
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {
                    "range": RANGE,
                    "values": [values_line],
                },
            ],
        },
    )

    try_request(request)


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
        date = ORIGIN_TIME + datetime.timedelta(days=current_date)
        calendar = get_col(row, headers_id["Calendar"])
        if not calendar:
            rich.print(
                f"Jumping event on {date.strftime('%Y-%m-%d')} since calendar is not defined"
            )
            continue

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

        calendar_id = calendars.get(calendar)
        if calendar_id is None:
            rich.print(f'Cannot find a calendar id associated to calendar "{calendar}"')
            raise KeyError

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
                rich.print(f'Deleted event "{get_col(row, headers_id["Activity"])}"')
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

                try_request(request)
                continue
            else:
                # There's something in the action cell, but not recognized
                rich.print(f"Unknown action {action}. Ignoring…")
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
        details = get_col(row, headers_id["Details"])
        add_to_gitlab = get_col(row, headers_id["Add Spent"])

        try:
            pid = projects[project]
        except KeyError:
            rich.print(
                f"Cannot find a project id, skipping comment to issue '{issue}'."
            )
            continue
        spent = get_col(row, headers_id["Spent"])
        if not spent:
            spent = 8
        try:
            url, gitlab_token = read_gitlab_token(config_dir)
        except ValueError as e:
            rich.print(f"ValueError: {e}")
            continue
        if add_to_gitlab == "":
            add_spent_time_on_gitlab_issue(
                url,
                gitlab_token,
                pid,
                issue.strip("#"),
                spent,
                details,
            )
            rich.print(f"Added {spent} hours to issue {project}{issue}")
        else:
            rich.print(f"Skipped reporting {spent} hours to issue {project}{issue}")


def read_gitlab_token(config_dir):
    if GITLAB_TOKEN_CONFIG not in os.listdir(config_dir):
        raise ValueError("No GitLab token configuration")
    token_file = config_dir / GITLAB_TOKEN_CONFIG
    with open(token_file) as f:
        info = json.load(f)
    return info["url"], info["token"]


def add_spent_time_on_gitlab_issue(
    gitlab_base_url, private_token, project_id, issue_id, spent, details
):
    # Authenticate to the GitLab API
    gl = gitlab.Gitlab(gitlab_base_url, private_token=private_token)

    # Get the issue object
    project = gl.projects.get(project_id)
    try:
        issue = project.issues.get(issue_id)
    except gitlab.GitlabGetError as e:
        rich.print(f"Invalid issue '{issue_id}'. Could not add time spent.")
        if "no gitlab" in details.lower():
            return
        raise RuntimeError(e)
    # Add a comment to the issue
    issue.notes.create({"body": f"/spend {spent}h"})


def get_calendars(sheet):
    RANGE = f"{get('CONTROLLER_SHEET_NAME', 'config')}!A2:B"
    calendars = (
        sheet.values()
        .get(spreadsheetId=get("CONTROLLER_SHEET_DOCUMENT_ID"), range=RANGE)
        .execute()
    )
    values = calendars.get("values", [])
    return {alias: id for [id, alias] in values}


def get_calendar_col_values(sheet, month, col_name):
    """Get all events ids for a month."""
    headers_ids = get_headers(sheet, month, indexes=True)
    col_of_interest = headers_ids.get(col_name)
    # transform a zero.based index to a capital letter
    col_of_interest = string.ascii_uppercase[col_of_interest]
    RANGE = f"{month}!{col_of_interest}2:{col_of_interest}"
    events = (
        sheet.values()
        .get(spreadsheetId=get("CONTROLLER_SHEET_DOCUMENT_ID"), range=RANGE)
        .execute()
    )
    values = events.get("values", [])
    return [e[0] for e in values if e]


def get_calendars_names(sheet, flat=True):
    """Get all calendars names, giving precedence to alias defined in column "linked_calendar".

    If multiple aliases are found, the first one will be used
    """
    RANGE = f"{get('CONTROLLER_SHEET_NAME', 'config')}!A2:C"
    calendars = (
        sheet.values()
        .get(spreadsheetId=get("CONTROLLER_SHEET_DOCUMENT_ID"), range=RANGE)
        .execute()
    )
    values = calendars.get("values", [])
    names = {}
    for cols in values:
        try:
            id, alias, linked_id = cols
        except ValueError:
            # no linked_id
            id, alias = cols
            linked_id = None
        if names.get(linked_id) or (names.get(id) and not linked_id):
            continue
        names[linked_id or id] = (
            alias if flat else {"alias": alias, "is_linked": bool(linked_id)}
        )
    return names


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
    creds = get_credentials(config_dir, SCOPES, "sheets-token.json")

    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()

    try:
        document_id = get("CONTROLLER_SHEET_DOCUMENT_ID")
    except KeyError:
        rich.print(
            "A value for CONTROLLER_SHEET_DOCUMENT_ID is required but "
            "is not specified in your ini file"
        )
        sys.exit(1)

    if month is None:
        sheets = sheet.get(spreadsheetId=document_id).execute()
        month = sheets["sheets"][-1]["properties"]["title"]
    rich.print("Sheet: {}".format(month))
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
