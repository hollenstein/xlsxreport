"""Command to compile a formatted Excel from a csv file and a formatting template."""
import os
import warnings

import click
import pandas as pd
import xlsxwriter

from xlsxreport import (
    get_template_path,
    TableSectionWriter,
    ReportTemplate,
    prepare_table_sections,
)


HELP = {
    "outfile": (
        "Name of the report file, by default the `INFILE` name is used with the file "
        "extension replaced by '.report.xlsx'."
    ),
    "outpath": (
        "Output path of the report file. If specified overrides the `outfile` option."
    ),
    "sep": "Delimiter to use for the input file, default is \\t.",
}


@click.command()
@click.argument("infile", type=click.Path(exists=True, readable=True))
@click.argument("template")
@click.option("--outfile", help=HELP["outfile"])
@click.option(
    "--outpath", help=HELP["outpath"], type=click.Path(exists=True, writable=True)
)
@click.option("--sep", default="\t", help=HELP["sep"])
def compile_excel_command(
    infile: str, template: str, outfile: str, outpath: str, sep: str
) -> None:
    """Create a formatted Excel report from a csv INFILE and a formatting TEMPLATE file.

    The TEMPLATE argument is first used to look for a file with the specified filepath.
    If no file is found, the XlsxReport appdata directory is searched for a file with
    the corresponding name.
    """
    report_path = _get_report_output_path(infile, outfile, outpath)
    template_path = get_template_path(template)
    if template_path is None:
        raise click.ClickException(
            f"Invalid value for `template`: '{template}' file not found."
        )

    click.echo(f"Generating formatted Excel report:")
    click.echo(f"----------------------------------")
    click.echo(f"Input file:    {infile}")
    click.echo(f"Template file: {template_path}")
    click.echo(f"Report file:   {report_path}")

    compile_excel(infile, template_path, report_path, sep)


def compile_excel(infile: str, template: str, outpath: str, sep: str = "\t") -> None:
    """Creates a formatted Excel report from a csv infile and a report template file.

    Args:
        infile: Path to the input csv file.
        template: Path to the formatting template file.
        outpath: Output path of the Excel report file.
        sep: Delimiter to use for the input file, by default \\t.
    """
    with warnings.catch_warnings():
        warnings.simplefilter(action="ignore", category=pd.errors.DtypeWarning)
        table = pd.read_csv(infile, sep=sep)

    report_template = ReportTemplate.load(template)
    table_sections = prepare_table_sections(report_template, table)
    with xlsxwriter.Workbook(outpath) as workbook:
        worksheet = workbook.add_worksheet("Report")
        section_writer = TableSectionWriter(workbook)
        section_writer.write_sections(
            worksheet, table_sections, settings=report_template.settings
        )


def _get_report_output_path(infile: str, outfile: str, outpath: str) -> str:
    """Return the path for the Excel report file."""
    if outpath:
        report_path = outpath
    elif outfile:
        report_path = os.path.join(os.path.dirname(infile), outfile)
    else:
        infilename = os.path.basename(infile)
        outfilename = ".".join(infilename.split(".")[:-1]) + ".report.xlsx"
        report_path = os.path.join(os.path.dirname(infile), outfilename)
    return report_path
