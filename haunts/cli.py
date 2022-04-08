"""Console script for haunts."""
import datetime
from email.policy import default
import os
import sys
from pathlib import Path

import click

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
    print(f"Manage you settings at {config.resolve()} and try again")
    sys.exit(0)
else:
    init(config)


@click.group()
@click.version_option(version=None, prog_name="haunts", message="Haunts %(version)s")
def haunts():
    pass


@click.command()
@click.argument("month")
@click.option(
    "--day",
    "-d",
    multiple=True,
)
def sync(month, day=[]):
    """Console script for haunts."""
    click.echo("Started calendars synchronization")
    init_calendars(config_dir)
    sync_report(
        config_dir,
        month,
        days=[datetime.datetime.strptime(d, "%Y-%m-%d") for d in day],
    )
    return 0


haunts.add_command(sync, "sync")


@click.command()
@click.argument("month")
@click.option(
    "--issue", "-i", default=None
)
@click.option(
    "--project", "-p", default=None
)
@click.option(
    "--explode", "-e", is_flag=True, default=False
)
@click.option(
    "-c1", default=report.COL_SIZES[0], show_default=True,
    help="Report first column size in characters."
)
@click.option(
    "-c2", default=report.COL_SIZES[1], show_default=True,
    help="Report second column size in characters."
)
@click.option(
    "-c3", default=report.COL_SIZES[2], show_default=True,
    help="Report third column size in characters."
)
@click.option(
    "-c4", default=report.COL_SIZES[3], show_default=True,
    help="Report fourth column size in characters (Detailed report)."
)
@click.option(
    "-c5", default=report.COL_SIZES[4], show_default=True,
    help="Report fifth column size in characters (Detailed report)."
)
def show_report(month, issue, project, explode, c1, c2, c3, c4, c5):
    """Shows number of spent hours for each issues and projects."""
    report.compute_report(
        config_dir=config_dir,
        month=month,
        issue=issue,
        project=project,
        explode=explode,
        col_sizes=[c1, c2, c3, c4, c5],
    )


haunts.add_command(show_report, "report")
