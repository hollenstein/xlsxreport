import os
import warnings

import click
import pandas as pd
import xlsxwriter

import xlsxreport.appdir
from xlsxreport.excel_writer import TableSectionWriter
from xlsxreport.template import ReportTemplate
from xlsxreport.compiler import prepare_table_sections


OUTFILE_DESCRIPTION = (
    "Name of the report file, by default the `INFILE` name is used with the file "
    "extension replaced by '.report.xlsx'."
)
OUTPATH_DESCRIPTION = (
    "Output path of the report file. If specified overrides the `outfile` option."
)
SEPARATOR_DESCRIPTION = "Delimiter to use for the input file, default is \\t."


@click.command()
@click.argument("infile", type=click.Path(exists=True))
@click.argument("template")
@click.option("--outfile", default="", help=OUTFILE_DESCRIPTION)
@click.option("--outpath", default="", help=OUTPATH_DESCRIPTION)
@click.option("--sep", default="\t", help=SEPARATOR_DESCRIPTION)
def report(infile: str, template: str, outfile: str, outpath: str, sep: str) -> None:
    """Create a formatted Excel report from csv INFILE and a formatting TEMPLATE file.

    The TEMPLATE argument is first used to look for a file with the specified filepath. If
    no file is found, the XlsxReport appdata directory is searched for a file with the
    corresponding name.
    """
    template_path = xlsxreport.appdir.get_template_path(template)
    if template_path is None:
        raise click.ClickException(
            f"Invalid value for `TEMPLATE`: '{template}' not found."
        )

    if outpath:
        if not os.path.isdir(os.path.dirname(outpath)):
            outdir = os.path.dirname(outpath)
            raise click.ClickException(
                f"Invalid value for `outpath`: '{outdir}' directory not found."
            )
        report_path = outpath
    elif outfile:
        report_path = os.path.join(os.path.dirname(infile), outfile)
    else:
        infilename = os.path.basename(infile)
        outfilename = ".".join(infilename.split(".")[:-1]) + ".report.xlsx"
        report_path = os.path.join(os.path.dirname(infile), outfilename)

    click.echo(f"Generating formatted Excel report:")
    click.echo(f"----------------------------------")
    click.echo(f"Input file:    {infile}")
    click.echo(f"Template file: {template_path}")

    with warnings.catch_warnings():
        warnings.simplefilter(action="ignore", category=pd.errors.DtypeWarning)
        table = pd.read_csv(infile, sep=sep)

    report_template = ReportTemplate.load(template_path)
    table_sections = prepare_table_sections(report_template, table)

    with xlsxwriter.Workbook(report_path) as workbook:
        worksheet = workbook.add_worksheet("Report")
        section_writer = TableSectionWriter(workbook)
        section_writer.write_sections(
            worksheet, table_sections, settings=report_template.settings
        )
    click.echo(f"Report file:   {report_path}")


@click.command()
def setup() -> None:
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


@click.group()
def cli():
    pass


cli.add_command(report)
cli.add_command(setup)


if __name__ == "__main__":
    cli()
