"""Command line interface for xlsxreport."""
import click

from xlsxreport.scripts.setup_appdir import setup_appdir_command
from xlsxreport.scripts.compile_excel import compile_excel_command


@click.group()
def cli():
    """Command line interface for xlsxreport."""
    pass


cli.add_command(setup_appdir_command, name="setup")
cli.add_command(compile_excel_command, name="compile")

if __name__ == "__main__":
    cli()
