"""Command to setup the xlsxreport app directory and copy default template files."""

import os

import click

import xlsxreport.appdir


HELP = {
    "setup": "Create XlsxReport app directory and copy default report template files.",
    "overwrite": (
        "Overwrite existing report template files when creating the app directory "
        "with the '--setup' option."
    ),
    "reveal": "Reveal the app directory in the file explorer.",
}


@click.command()
@click.option("--setup", is_flag=True, default=False, help=HELP["setup"])
@click.option("--overwrite", is_flag=True, default=False, help=HELP["overwrite"])
@click.option("--reveal", is_flag=True, default=False, help=HELP["reveal"])
def appdir_command(setup, overwrite, reveal) -> None:
    """Locate app directory, optionally create the directory and copy default report
    template files."""
    appdir = xlsxreport.appdir.locate_appdir()
    if not setup:
        if os.path.isdir(appdir):
            click.echo(appdir)
        else:
            click.echo(
                "XlsxReport app directory not found, run `xlsxreport appdir --setup` "
                "to create the app directory."
            )
    if setup:
        if not os.path.isdir(appdir):
            click.echo(f"Creating XlsxReport app directory at: {appdir}")
        else:
            click.echo(f"XlsxReport app directory found at: {appdir}")

        if overwrite:
            click.echo(
                "Copying default report templates to the app directory, overwriting "
                "existing files."
            )
        else:
            click.echo("Copying missing default report templates to the app directory.")
        xlsxreport.appdir.setup_appdir(overwrite_templates=overwrite)
    if reveal:
        click.launch(appdir)
