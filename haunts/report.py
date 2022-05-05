import sys
import datetime

from googleapiclient.discovery import build

from .ini import get
from . import spreadsheet
from .calendars import ORIGIN_TIME


FULL_EVENT_HOURS = 8
ROW_FORMAT = "{:<{a}} {:<{b}} {:<{c}}"
SEP_FORMAT = "{:-^{a}} {:-^{b}} {:-^{c}}"
ROW_FORMAT_MORE = "{:<{a}} {:<{b}} {:<{c}} {:<{d}} {:>{e}}"
SEP_FORMAT_MORE = "{:-^{a}} {:-^{b}} {:-^{c}} {:-^{d}} {:-^{e}}"
COL_SIZES = (20, 20, 5, 10, 30)


def compute_hours_report(month):
    service = build("sheets", "v4", credentials=spreadsheet.creds)
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

    data = (
        sheet.values()
        .get(
            spreadsheetId=document_id,
            range=f"{month}!A2:ZZ",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
    )

    headers_id = spreadsheet.get_headers(sheet, month, indexes=True)

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
        report.setdefault((project, issue), []).append(
            {"date": date, "time": float(spent), "title": title})

    return report


def print_report(report, c1=20, c2=20, c3=5):
    print(ROW_FORMAT.format("Project", "Issue", "Hours", a=c1, b=c2, c=c3))
    print(SEP_FORMAT.format(*[""] * 3, a=c1, b=c2, c=c3))
    for pair, values in report.items():
        print(ROW_FORMAT.format(
            *pair, sum(v["time"] for v in values), a=c1, b=c2, c=c3)
        )


def print_detailed_report(report, c1=20, c2=20, c3=5, c4=10, c5=30):
    print(ROW_FORMAT_MORE.format(
        "Project", "Issue", "Hours", "Date", "Title",
        a=c1, b=c2, c=c3, d=c4, e=c5)
    )
    print(SEP_FORMAT_MORE.format(
        *[""] * 5, a=c1, b=c2, c=c3, d=c4, e=c5)
    )
    for pair, values in report.items():
        for v in values:
            print(ROW_FORMAT_MORE.format(
                *pair, v["time"], v["date"].strftime("%d/%m/%Y"), v["title"],
                a=c1, b=c2, c=c3, d=c4, e=c5)
            )


def compute_report(
    config_dir, month, issue, project, explode, col_sizes
):
    c1, c2, c3, c4, c5 = col_sizes
    spreadsheet.get_credentials(config_dir)
    report = compute_hours_report(month)
    if issue is not None:
        report = {pair: report[pair] for pair in report if issue in pair[1]}
    if project is not None:
        report = {pair: report[pair] for pair in report if project in pair[0]}

    if explode:
        print_detailed_report(report, c1, c2, c3, c4, c5)
    else:
        print_report(report, c1, c2, c3)
