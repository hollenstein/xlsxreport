"""Command line interface for xlsxreport."""
import click

from xlsxreport.scripts.setup_appdir import setup_appdir
from xlsxreport.scripts.compile_excel import compile_excel


@click.group()
def cli():
    """Command line interface for xlsxreport."""
    pass


cli.add_command(setup_appdir, name="setup")
cli.add_command(compile_excel, name="compile")


if __name__ == "__main__":
    cli()
