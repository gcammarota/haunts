"""Console script for haunts."""
import datetime
import os
import sys
from pathlib import Path

import click

from .ini import create_default, init

from .calendars import init as init_calendars
from .spreadsheet import sync_report, compute_report

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
    "--day",
    "-d",
    multiple=True,
)
def report(month, day=[]):
    """Shows number of spent hours for each issues."""
    click.echo("Issues report [hours]")
    compute_report(
        config_dir,
        month,
        days=[datetime.datetime.strptime(d, "%Y-%m-%d") for d in day]
    )


haunts.add_command(report, "report")


if __name__ == "__main__":
    sys.exit(sync())  # pragma: no cover
