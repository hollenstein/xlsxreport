"""Command line interface for xlsxreport."""
import click

from xlsxreport.scripts.appdir import appdir_command
from xlsxreport.scripts.compile_excel import compile_excel_command


@click.group()
def cli():
    """Command line interface for xlsxreport."""
    pass


cli.add_command(appdir_command, name="appdir")
cli.add_command(compile_excel_command, name="compile")

if __name__ == "__main__":
    cli()
