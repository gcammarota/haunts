import sys
import datetime

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
    print("Month: {}".format(month))

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
        project = spreadsheet.get_col(row, headers_id["Project"])
        action = spreadsheet.get_col(row, headers_id["Action"])
        if not spent:
            spent = FULL_EVENT_HOURS
        report.setdefault((project, issue), []).append(
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
        project = spreadsheet.get_col(row, headers_id["Project"])
        if not spent:
            spent = FULL_EVENT_HOURS
        report.setdefault(date, []).append(
            {"project": project, "issue": issue, "time": float(spent), "title": title})

    # Select only the days with total spent time less than 8h
    missing_report = {}
    for d, r in report.items():
        day_total_spent = sum([el["time"] for el in r])
        if day_total_spent != 8:
            missing_report.update({(d, day_total_spent): r})

    return missing_report


def tune_report(config_dir, month=None, issue=None, project=None):
    spreadsheet.get_credentials(config_dir)
    report = compute_report(month)
    if issue is not None:
        report = {pair: report[pair] for pair in report if issue in pair[1]}
    if project is not None:
        report = {pair: report[pair] for pair in report if project in pair[0]}
    return report


def print_table_header(row_format, sep_format, *args, **kwargs):
    col_rows = len(args)
    print(sep_format.format(*[""] * col_rows, **kwargs))
    print(row_format.format(*args, **kwargs))
    print(sep_format.format(*[""] * col_rows, **kwargs))


def print_report(config_dir, month, issue, project, unreported_only, col_sizes):
    report = tune_report(config_dir, month, issue, project)

    c1, c2, c3 = col_sizes
    print_table_header(
        ROW_FORMAT, SEP_FORMAT,
        "Project", "Issue", "Added", "Losts",
        a=c1, b=c2, c=c3
    )
    for pair, values in report.items():
        unreported = sum(v["time"] for v in values if not v["action"])
        reported = sum(v["time"] for v in values if v["action"] == "I")
        if unreported_only and unreported == 0:
            continue  # show only unreported hours
        print(ROW_FORMAT.format(
            *pair, reported, unreported, a=c1, b=c2, c=c3)
        )


def print_detailed_report(config_dir, month, issue, project, unreported_only, col_sizes):
    report = tune_report(config_dir, month, issue, project)

    c1, c2, c3, c4, c5 = col_sizes
    print_table_header(
        ROW_FORMAT_MORE, SEP_FORMAT_MORE,
        "Project", "Issue", "Time", "Added", "Date", "Title",
        a=c1, b=c2, c=c3, d=c4, e=c5
    )
    for pair, values in report.items():
        for v in values:
            is_reported = v["action"] == "I"
            if unreported_only and is_reported:
                continue  # show only unreported issues
            print(ROW_FORMAT_MORE.format(
                *pair, v["time"], f"{is_reported}", v["date"].strftime("%d/%m/%Y").strip(), v["title"],
                a=c1, b=c2, c=c3, d=c4, e=c5)
            )


def print_mail(config_dir, month, issue, project):
    spreadsheet.get_credentials(config_dir)
    document, document_id = get_document()
    if month is None:
        month = get_month(document, document_id)
    report = tune_report(config_dir, month, issue, project)
    m, y = month.split(" ")
    print("")
    print(f"Presenze {month}\n\n")
    print("Ciao,\n")
    print(f"a {m} sempre presente", end="")
    if report:
        print(" tranne il:\n")
    else:
        print(".")
    for pair, values in report.items():
        for v in values:
            print(f"- {v['date'].strftime('%d').strip()}: {v['title']}")
    print("\n\nCiao,\n")


def print_incomplete_days(config_dir, month=None):
    spreadsheet.get_credentials(config_dir)
    report = compute_missing(month)
    if not report:
        print("🍺 Everything is fine!🍺 ")
        return
    print_table_header(
        INCOMPLETE_FORMAT, INCOMPLETE_SEP_FORMAT,
        "Incomplete", "Hours",
        a=COL_SIZES[3], b=COL_SIZES[2]
    )
    for day, partial in report:
        partial = RED.format(str(partial)) if partial < FULL_EVENT_HOURS \
            else GREEN.format(str(partial))
        print(INCOMPLETE_FORMAT.format(
            day.strftime("%d/%m/%Y"), partial,
            a=COL_SIZES[3], b=COL_SIZES[2])
        )
