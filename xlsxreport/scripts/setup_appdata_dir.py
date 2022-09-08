import os
import warnings

import click

import xlsxreport


@click.command()
def cli() -> None:
    """Description of the XlsxReport function."""

    data_dir = xlsxreport.locate_data_dir()

    click.echo(f"Creating 'XlsxReport' folder in the local user data directory at:")
    click.echo(f"\t{data_dir}")
    xlsxreport.setup_data_dir()
    click.echo(f"The 'XlsxReport' folder was successfully created.")
