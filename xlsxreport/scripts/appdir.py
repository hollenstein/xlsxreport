"""Command to setup the xlsxreport app directory and copy default template files."""

import os

import click
import pathlib

import xlsxreport.appdir


HELP = {
    "setup": "Create XlsxReport app directory and copy default table template files.",
    "overwrite": (
        "Overwrite existing table template files when creating the app directory "
        "with the '--setup' option."
    ),
    "reveal": "Reveal the app directory in the file explorer.",
    "templates": "List the table template files in the app directory.",
}
PATH_COLOR = "yellow"


@click.command()
@click.option("--setup", is_flag=True, default=False, help=HELP["setup"])
@click.option("--overwrite", is_flag=True, default=False, help=HELP["overwrite"])
@click.option("-r", "--reveal", is_flag=True, default=False, help=HELP["reveal"])
@click.option("-t", "--templates", is_flag=True, default=False, help=HELP["reveal"])
def appdir_command(setup, overwrite, reveal, templates) -> None:
    """Locate app directory, optionally create the directory and copy default table
    template files."""
    appdir = xlsxreport.appdir.locate_appdir()
    if setup:
        if not os.path.isdir(appdir):
            click.echo(
                f"Creating XlsxReport app directory at: {_format_filename(appdir)}"
            )
        else:
            click.echo(f"XlsxReport app directory found at: {_format_filename(appdir)}")

        if overwrite:
            click.echo(
                "Copying default table templates to the app directory, overwriting "
                "existing files"
            )
        else:
            click.echo("Copying missing default table templates to the app directory")
        xlsxreport.appdir.setup_appdir(overwrite_templates=overwrite)
    elif not os.path.isdir(appdir):
        click.echo(
            "XlsxReport app directory not found, run `xlsxreport appdir --setup` "
            "to create the app directory."
        )
        return

    if not any([setup, reveal, templates]):
        click.echo(_format_filename(appdir))
    if reveal:
        click.launch(appdir)
    if templates:
        click.echo("Table template files in the app directory:")
        for template in xlsxreport.appdir.get_appdir_templates():
            click.echo(f"  {_format_filename(template)}")


def _format_filename(filename: str) -> str:
    return click.style(
        pathlib.Path(click.format_filename(filename)).as_posix(), fg=PATH_COLOR
    )
