import sys
import datetime
import locale

from googleapiclient.discovery import build

from .ini import get
from . import spreadsheet
from .calendars import ORIGIN_TIME


FULL_EVENT_HOURS = 8
ROW_FORMAT = "{:<{a}} {:<{b}} {:>{c}} {:>{c}}"
SEP_FORMAT = "{:-^{a}} {:-^{b}} {:-^{c}} {:-^{c}}"
ROW_FORMAT_MORE = "{:<{a}} {:<{b}} {:<{c}} {:<{c}} {:<{d}} {:>{e}}"
SEP_FORMAT_MORE = "{:-^{a}} {:-^{b}} {:-^{c}} {:-^{c}} {:-^{d}} {:-^{e}}"
INCOMPLETE_FORMAT = "{:<{a}} {:^{b}}"
INCOMPLETE_SEP_FORMAT = "{:-^{a}} {:-^{b}}"
COL_SIZES = (20, 20, 5, 10, 30)
RED = "\x1b[1;31;40m{}\x1b[0m"
GREEN = "\x1b[1;32;40m{}\x1b[0m"


def get_document():
    service = build("sheets", "v4", credentials=spreadsheet.creds)
    # Call the Sheets API
    document = service.spreadsheets()
    try:
        document_id = get("CONTROLLER_SHEET_DOCUMENT_ID")
    except KeyError:
        print(
            "A value for CONTROLLER_SHEET_DOCUMENT_ID is required but "
            "is not specified in your ini file"
        )
        sys.exit(1)
    return document, document_id


def get_month(document, document_id):
    sheets = document.get(spreadsheetId=document_id).execute()
    return sheets["sheets"][-1]["properties"]["title"]


def prepare_report_data(month=None):
    document, document_id = get_document()
    if month is None:
        month = get_month(document, document_id)
    print("Sheet: {}\n".format(month))

    data = (
        document.values()
        .get(
            spreadsheetId=document_id,
            range=f"{month}!A2:ZZ",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
    )

    headers_id = spreadsheet.get_headers(document, month, indexes=True)
    return data, headers_id


def compute_report(month=None):
    data, headers_id = prepare_report_data(month)

    report = {}
    for row in data["values"]:
        date = spreadsheet.get_col(row, headers_id["Date"])
        if not date:
            continue
        date = ORIGIN_TIME + datetime.timedelta(days=date)
        issue = spreadsheet.get_col(row, headers_id["Issue"])
        title = spreadsheet.get_col(row, headers_id["Title"])
        spent = spreadsheet.get_col(row, headers_id["Spent"])
        calendar = spreadsheet.get_col(row, headers_id["Calendar"])
        project = spreadsheet.get_col(row, headers_id["Project"])
        action = spreadsheet.get_col(row, headers_id["Action"])
        if not spent:
            spent = FULL_EVENT_HOURS
        report.setdefault((calendar, project, issue), []).append(
            {"date": date, "time": float(spent), "title": title, "action": action})

    return report


def compute_missing(month=None):
    data, headers_id = prepare_report_data(month)

    report = {}
    for row in data["values"]:
        date = spreadsheet.get_col(row, headers_id["Date"])
        if not date:
            continue
        date = ORIGIN_TIME + datetime.timedelta(days=date)
        issue = spreadsheet.get_col(row, headers_id["Issue"])
        title = spreadsheet.get_col(row, headers_id["Title"])
        spent = spreadsheet.get_col(row, headers_id["Spent"])
        calendar = spreadsheet.get_col(row, headers_id["Calendar"])
        project = spreadsheet.get_col(row, headers_id["Project"])
        if not spent:
            spent = FULL_EVENT_HOURS
        report.setdefault(date, []).append(
            {"calendar": calendar, "project": project, "issue": issue, "time": float(spent), "title": title})

    # Select only the days with total spent time less than 8h
    missing_report = {}
    for d, r in report.items():
        day_total_spent = sum([el["time"] for el in r])
        if day_total_spent != 8:
            missing_report.update({(d, day_total_spent): r})

    return missing_report


def tune_report(report, issue=None, project=None, calendar=None):
    if issue is not None:
        report = {triplet: report[triplet] for triplet in report if issue in triplet[2]}
    if project is not None:
        report = {triplet: report[triplet] for triplet in report if project in triplet[1]}
    if calendar is not None:
        report = {triplet: report[triplet] for triplet in report if calendar in triplet[0]}
    return report


def print_table_header(row_format, sep_format, *args, **kwargs):
    col_rows = len(args)
    print(sep_format.format(*[""] * col_rows, **kwargs))
    print(row_format.format(*args, **kwargs))
    print(sep_format.format(*[""] * col_rows, **kwargs))


def prepare_report(config_dir, month=None, issue=None, project=None, calendar=None):
    spreadsheet.get_credentials(config_dir)
    report = compute_report(month)
    return tune_report(report, issue, project, calendar)


def prepare_mail(config_dir, month):
    spreadsheet.get_credentials(config_dir)
    if month is None:
        month = get_month(*get_document())
    report = compute_report(month)
    tuned_report = tune_report(report, calendar="ferie")
    prep = "ad" if month.startswith("A") else "a"
    new_line = "\n"
    holidays = ""
    if tuned_report:
        holidays += " tranne il:\n"
        for _, values in tuned_report.items():
            for v in values:
                holidays += f"\n- {v['date'].strftime('%d').strip()}: {v['title']}"
    mail = (
        f"Ciao,"
        f"{new_line * 2}"
        f"{prep} {month.split(' ')[0]} sempre presente"
        f"{holidays}"
        f"{new_line * 3}"
        f"Ciao,"
        f"{new_line}"
    )
    return mail


def prepare_incomplete_days(config_dir, month=None):
    spreadsheet.get_credentials(config_dir)
    return compute_missing(month)
