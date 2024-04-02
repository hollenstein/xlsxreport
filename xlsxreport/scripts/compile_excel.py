"""Command to compile a formatted Excel from a csv file and a formatting template."""

import os
import warnings

import click
import pandas as pd
import xlsxwriter  # type: ignore

from xlsxreport import (
    get_template_path,
    SectionWriter,
    TableTemplate,
    prepare_compiled_sections,
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
    "reveal": "Open the compiled Excel report file in the default application.",
}


@click.command()
@click.argument("infile", type=click.Path(exists=True, readable=True))
@click.argument("template")
@click.option("-o", "--outfile", help=HELP["outfile"])
@click.option(
    "--outpath", type=click.Path(exists=True, writable=True), help=HELP["outpath"]
)
@click.option("-s", "--sep", default="\t", help=HELP["sep"])
@click.option("-r", "--reveal", is_flag=True, default=False, help=HELP["reveal"])
def compile_excel_command(infile, template, outfile, outpath, sep, reveal) -> None:
    """Create a formatted Excel report from a csv INFILE and a formatting TEMPLATE file.

    The TEMPLATE argument is first used to look for a file with the specified filepath.
    If no file is found, the xlsxreport appdata directory is searched for a file with
    the corresponding name.
    """
    report_path = _get_report_output_path(infile, outfile, outpath)
    template_path = get_template_path(template)
    if template_path is None:
        raise click.ClickException(
            f"Invalid value for 'TEMPLATE': Path '{click.format_filename(template)}' "
            "does not exist."
        )

    click.echo(f"Generating formatted Excel report:")
    click.echo(f"----------------------------------")
    click.echo(f"Input file:    {click.format_filename(infile)}")
    click.echo(f"Template file: {click.format_filename(template_path)}")
    click.echo(f"Report file:   {click.format_filename(report_path)}")

    compile_excel(infile, template_path, report_path, sep)
    if reveal:
        click.launch(report_path)


def compile_excel(infile: str, template: str, outpath: str, sep: str = "\t") -> None:
    """Creates a formatted Excel report from a csv infile and a table template file.

    Args:
        infile: Path to the input csv file.
        template: Path to the formatting template file.
        outpath: Output path of the Excel report file.
        sep: Delimiter to use for the input file, by default \\t.
    """
    with warnings.catch_warnings():
        warnings.simplefilter(action="ignore", category=pd.errors.DtypeWarning)
        table = pd.read_csv(infile, sep=sep)

    table_template = TableTemplate.load(template)
    compiled_sections = prepare_compiled_sections(table_template, table)
    with xlsxwriter.Workbook(outpath) as workbook:
        worksheet = workbook.add_worksheet("Report")
        section_writer = SectionWriter(workbook)
        section_writer.write_sections(
            worksheet, compiled_sections, settings=table_template.settings
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
