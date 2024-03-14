"""Console script for haunts."""
import datetime
import os
import sys
from pathlib import Path
from typing import List

import rich
import typer

from .ini import create_default, init

from .calendars import init as init_calendars
from .spreadsheet import sync_report
from . import report

config_dir = Path(os.path.expanduser("~/.haunts"))

if not config_dir.is_dir():
    print(f"Creating config directory at {config_dir.resolve()}")
    config_dir.mkdir()
    print("…created")


config = Path(os.path.expanduser(f"{config_dir.resolve()}/haunts.ini"))
if not config.is_file():
    create_default(config)
    print(f"Manage your settings at {config.resolve()} and try again")
    sys.exit(0)
else:
    init(config)

app = typer.Typer(
    no_args_is_help=True,
    help="Haunts CLI",
)


def haunts() -> None:
    app()


@app.command()
def push(
    month: str = typer.Option(None, "-m", "--month"),
    days: List[str] = typer.Option([], "-d", "-days")
):
    """Create events on calendar for the month or specific days"""
    rich.print("Started calendars synchronization")
    init_calendars(config_dir)
    sync_report(
        config_dir,
        month,
        days=[datetime.datetime.strptime(d, "%Y-%m-%d") for d in days],
    )


@app.command("check")
def check_spent_hours(
    month: str = typer.Option(None, "-m", "--month"),
):
    """Show number of spent hours for each anomalous (!= 8) worked day."""
    res = report.prepare_incomplete_days(config_dir=config_dir, month=month)
    if not res:
        rich.print("🍺 Everything is fine!🍺 ")
    else:
        table = rich.table.Table("Incomplete", "Hours")
        for day, partial in res:
            if partial < report.FULL_EVENT_HOURS:
                partial = f"[red]{str(partial)}[/]"
            else:
                partial = f"[green]{str(partial)}[/]"
            table.add_row(
                day.strftime("%d/%m/%Y"),
                partial
            )
        rich.print(table)


@app.command(name="report")
def show_report(
    month: str = typer.Option(None, "-m", "--month"),
    issue: str = typer.Option(None, "-i", "--issue"),
    project: str = typer.Option(None, "-p", "--project"),
    calendar: str = typer.Option(None, "-c", "--calendar"),
):
    """Show number of spent hours for each issues and projects aggregated."""
    table = rich.table.Table("Calendar", "Project", "Issue", "Added", "Losts")
    res = report.prepare_report(
        config_dir=config_dir,
        month=month,
        issue=issue,
        project=project,
        calendar=calendar,
    )
    for triplet, values in res.items():
        table.add_row(
            triplet[0],
            triplet[1],
            triplet[2],
            str(sum(v["time"] for v in values if v["action"] == "I")),
            str(sum(v["time"] for v in values if not v["action"])),
        )
    rich.print(table)


@app.command(name="report-full")
def show_detailed_report(
    month: str = typer.Option(None, "-m", "--month"),
    issue: str = typer.Option(None, "-i", "--issue"),
    project: str = typer.Option(None, "-p", "--project"),
    calendar: str = typer.Option(None, "-c", "--calendar"),
):
    """Show number of spent hours for each issues and projects splitted."""
    table = rich.table.Table("Calendar", "Project", "Issue", "Time", "Added", "Date", "Title")
    res = report.prepare_report(
        config_dir=config_dir,
        month=month,
        issue=issue,
        project=project,
        calendar=calendar,
    )
    for triplet, values in res.items():
        for v in values:
            table.add_row(
                triplet[0],
                triplet[1],
                triplet[2],
                str(v["time"]),
                str(v["action"] == "I"),
                v["date"].strftime("%d/%m/%Y").strip(),
                v["title"],
            )
    rich.print(table)


@app.command()
def mail(month: str = typer.Option(None, "-m", "--month")):
    """Print a reporting mail with a recap of the holidays enjoyed in the last
    (in the spreadsheet file) or selected month

    """
    text = report.prepare_mail(
        config_dir=config_dir,
        month=month,
    )
    rich.print(text)
