import os
import warnings

import click
import pandas as pd

import xlsxreport


@click.command()
@click.argument("infile")
@click.argument("config")
@click.option(
    "--outfile",
    default="report.xlsx",
    help="Name of the report file, default is 'report.xlsx'.",
)
@click.option(
    "--outpath",
    default=None,
    help=(
        "Output path of the report file, if specified overrides the outfile parameter."
    ),
)
@click.option(
    "--sep", default="\t", help="Delimiter to use for 'infile', default is \\t."
)
def cli(infile: str, config: str, outfile: str, outpath: str, sep: str) -> None:
    """Description of the XlsxReport function."""
    if outpath is not None:
        report_path = outpath
    else:
        report_path = os.path.join(os.path.dirname(infile), outfile)

    if os.path.isfile(config):
        config_path = config
    else:
        config_path = xlsxreport.get_config_file(config)

    if not os.path.isfile(config_path):
        click.echo("Config file not found: %s" % config_path)

    click.echo(f"\n")
    click.echo(f"Preparing xlsx report:")
    click.echo(f"----------------------")
    click.echo(f"\tFile:   {infile}")
    click.echo(f"\tConfig: {config_path}\n")

    with warnings.catch_warnings():
        warnings.simplefilter(action="ignore", category=pd.errors.DtypeWarning)
        table = pd.read_csv(infile, sep=sep)

    with xlsxreport.Reportbook(report_path) as reportbook:
        protein_sheet = reportbook.add_datasheet("Proteins")
        protein_sheet.apply_configuration(config_path)
        protein_sheet.add_data(table)
        protein_sheet.write_data()
    click.echo(f"\tReport written to:")
    click.echo(f"\t{report_path}")


if __name__ == "__main__":
    cli()
