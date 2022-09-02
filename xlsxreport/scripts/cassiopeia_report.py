import os
import warnings

import click
import pandas as pd

import xlsxreport


@click.command()
@click.argument("infile")
@click.option(
    "--config",
    default="cassiopeia.yaml",
    help="Name of the config file, default is 'cassiopeia.yaml'.",
)
@click.option(
    "--outfile",
    default="cassiopeia_report.xlsx",
    help="Name of the report file, default is 'cassiopeia_report.xlsx'.",
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

    # Replace comparison column names
    replace_column_tags = [
        ("P.Value_", "P.Value"),
        ("adj.P.Val_", "adj.P.Val"),
        ("logFC_", "logFC"),
        ("AveExpr_", "AveExpr"),
    ]
    for old_tag, new_tag in replace_column_tags:
        table.columns = [c.replace(old_tag, new_tag) for c in table.columns]

    # Sort rows
    try:
        sort_rows_by = "MS/MS count"
        sort_ascending = [False]
        table.sort_values(
            sort_rows_by,
            ascending=sort_ascending,
            na_position="first",
            inplace=True,
        )
    except KeyError:
        pass

    # Generate excel report
    with xlsxreport.Reportbook(report_path) as reportbook:
        protein_sheet = reportbook.add_datasheet("Proteins")
        protein_sheet.apply_configuration(config_path)
        protein_sheet.add_data(table)
        protein_sheet.write_data()
    click.echo(f"\tReport written to:")
    click.echo(f"\t{report_path}")


if __name__ == "__main__":
    cli()
