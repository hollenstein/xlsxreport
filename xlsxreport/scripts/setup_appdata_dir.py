import os
import warnings

import click

import xlsxreport


@click.command()
def cli() -> None:
    """Description of the XlsxReport function."""

    data_dir = xlsxreport.locate_data_dir()
    if os.path.isdir(data_dir):
        click.echo(f"App data directory for XlsxReport found at:")
        click.echo(f"\t{data_dir}")
    else:
        click.echo(f"Creating XlsxReport folder in the local user data directory at:")
        click.echo(f"\t{data_dir}")
    click.echo("Copying default XlsxReport config files to the app data directory ...")
    xlsxreport.setup_data_dir()
    click.echo(f"\tConfig files were successfully copied.")
