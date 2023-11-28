"""Command to setup the xlsxreport app directory and copy default template files."""
import os

import click

import xlsxreport.appdir


@click.command()
def setup_appdir() -> None:
    """Setup app directory and copy default template files."""
    data_dir = xlsxreport.appdir.locate_appdir()
    if os.path.isdir(data_dir):
        click.echo(f"App data directory for XlsxReport found at:")
        click.echo(f"  {data_dir}")
    else:
        click.echo(f"Creating XlsxReport folder in the local user data directory at:")
        click.echo(f"  {data_dir}")
    click.echo(
        "Copying default XlsxReport template files to the app data directory ..."
    )
    xlsxreport.appdir.setup_appdir(overwrite_templates=True)
    click.echo(f"  Template files were successfully copied.")
