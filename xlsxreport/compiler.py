from dataclasses import dataclass, field
from typing import Iterable, Optional
import pandas as pd
from xlsxreport.template import ReportTemplate


DEFAULT_COL_WIDTH = 64
DEFAULT_FORMAT = {"num_format": "@"}


@dataclass
class TableSection:
    """Contains information for writing and formatting a section of a table."""

    data: pd.DataFrame
    column_formats: dict = field(default_factory=dict)
    column_conditionals: dict = field(default_factory=dict)
    column_widths: dict = field(default_factory=dict)
    header_formats: dict = field(default_factory=dict)
    supheader: str = ""
    supheader_format: dict = field(default_factory=dict)
    section_conditional: str = ""


# Missing from the compile_table_sections:
# 1) test which type the section template is
# 2) use different functions according to the type of the section template
# 3) Apply section type specific data manipulations
#    - remove tag from column names
#    - replace comparison tag in column names
#    - log2 transformation of column values
# 4) Apply common data manipulations
#    - Replace missing values / NaNs
# 5) Manipulation of formats
#    - Add borders to certain column formats if it was specified in the template


def compile_table_sections(
    report_template: ReportTemplate, table: pd.DataFrame
) -> list[TableSection]:
    """Compile table sections from a report template and a table."""

    table_sections = []
    for section_template in report_template.sections.values():
        selected_cols = eval_section_columns(section_template, table.columns)
        col_formats = eval_column_formats(
            selected_cols, section_template, report_template.formats, DEFAULT_FORMAT
        )
        col_conditionals = eval_column_conditional_formats(
            selected_cols, section_template, report_template.conditional_formats
        )
        default_width = report_template.settings.get("column_width", DEFAULT_COL_WIDTH)
        col_widths = eval_column_widths(selected_cols, section_template, default_width)
        header_formats = eval_header_formats(
            selected_cols, section_template, report_template.formats
        )
        supheader = section_template.get("supheader", "")
        supheader_format = eval_supheader_format(
            section_template, report_template.formats
        )
        section_conditional = eval_section_conditional_format(
            section_template, report_template.conditional_formats
        )

        table_section = TableSection(
            table[selected_cols].copy(),
            column_formats=col_formats,
            column_conditionals=col_conditionals,
            column_widths=col_widths,
            header_formats=header_formats,
            supheader=supheader,
            supheader_format=supheader_format,
            section_conditional=section_conditional,
        )
        table_sections.append(table_section)
    return table_sections


def eval_section_columns(template_section: dict, columns: Iterable[str]) -> list[str]:
    """Select columns from the template that are present in the table.

    Args:
        template_section: A dictionary containing the columns to be selected as the
            values of the "columns" key.
        columns: A list of column names to select from.

    Returns:
        A list of column names that are present in both the template and the table.
    """
    selected_columns = [col for col in template_section["columns"] if col in columns]
    return selected_columns


def eval_column_formats(
    columns: str,
    section_template: dict,
    format_templates: dict,
    default_format: Optional[dict] = None,
) -> dict:
    """Returns format descriptions for each column in the section.

    Args:
        columns: A list of column names.
        section_template: A dictionary containing the format names for columns.
        format_templates: A dictionary containing the format descriptions for each
            format name.
        default_format: Optional, the format description to use if no general format and
            no column format is specified in the `format_templates`.

    Returns:
        A dictionary containing format descriptions for each column.
    """
    default_format = {} if default_format is None else default_format
    section_format = section_template.get("format", None)
    column_formats = {}
    for col in columns:
        format_name = section_format
        if "column_format" in section_template:
            format_name = section_template["column_format"].get(col, section_format)

        column_formats[col] = format_templates.get(format_name, default_format).copy()
    return column_formats


def eval_column_conditional_formats(
    columns: str,
    section_template: dict,
    format_templates: dict,
) -> dict:
    """Returns conditional format descriptions for each column in the section.

    Args:
        columns: A list of column names.
        section_template: A dictionary containing the conditional format names for
            columns.
        format_templates: A dictionary containing the conditional format descriptions
            for each format name. If a format name is not present in the
            `format_templates`, an empty dictionary is used instead.

    Returns:
        A dictionary containing conditional format descriptions for each column.
    """
    default_format = {}
    column_formats = {}
    for col in columns:
        format_name = None
        if "column_conditional" in section_template:
            format_name = section_template["column_conditional"].get(col, None)
        column_formats[col] = format_templates.get(format_name, default_format).copy()
    return column_formats


def eval_column_widths(
    columns: str,
    section_template: dict,
    default_width: int = 64,
) -> dict:
    """Returns column widths for each column in the section.

    Args:
        columns: A list of column names.
        section_template: A dictionary containing the column widths for columns.
        default_width: The default column width to use if no column width is specified
            in the `section_template`.

    Returns:
        A dictionary containing column widths for each column.
    """
    column_widths = {}
    for col in columns:
        column_widths[col] = section_template.get("width", default_width)
    return column_widths


def eval_header_formats(
    columns: str, section_template: dict, format_templates: dict
) -> dict:
    """Returns format descriptions for each column header in the section.

    Header format descriptions defined in the `section_template` update the one from the
    `format_templates`.

    Args:
        columns: A list of column names.
        section_template: A dictionary that can contain a "header_format" description.
        format_templates: A dictionary that can contain a "header" format description.

    Returns:
        A dictionary containing header format descriptions for each column.
    """
    temmplate_format = format_templates.get("header", {})
    section_format = section_template.get("header_format", {})
    header_format = dict(temmplate_format, **section_format)
    column_header_formats = {col: header_format.copy() for col in columns}
    return column_header_formats


def eval_supheader_format(section_template: dict, format_templates: dict) -> dict:
    """Returns a format descriptions for the supheader.

    Supheader format description defined in the `section_template` updates the one from
    the `format_templates`.

    Args:
        columns: A list of column names.
        section_template: A dictionary that can contain a "supheader_format"
            description.
        format_templates: A dictionary that can contain a "supheader" format description.

    Returns:
        A dictionary describing the supheader format.
    """
    temmplate_format = format_templates.get("supheader", {})
    section_format = section_template.get("supheader_format", {})
    supheader_format = dict(temmplate_format, **section_format)
    return supheader_format


def eval_section_conditional_format(
    section_template: dict, format_templates: dict
) -> dict:
    """Returns a conditional format description of a section.

    Args:
        section_template: A dictionary that can contain a conditional format name with
            the key "conditional".
        format_templates: A dictionary containing the conditional format descriptions
            for each conditional format name. If a format name is not present in the
            `format_templates`, an empty dictionary is used instead.

    Returns:
        A dictionary containing a conditional format description.
    """
    section_format_name = section_template.get("conditional", None)
    section_conditional = format_templates.get(section_format_name, {}).copy()
    return section_conditional
