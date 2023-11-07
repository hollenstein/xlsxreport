from dataclasses import dataclass, field
from enum import Enum
from typing import Iterable, Optional, Protocol

import numpy as np
import pandas as pd

from xlsxreport.template import ReportTemplate


BORDER_TYPE: int = 2  # 2 = thick line, see xlsxwriter.format.Format().set_border()
DEFAULT_COL_WIDTH: float = 64
DEFAULT_FORMAT: dict = {"num_format": "@"}
NAN_REPLACEMENT_SYMBOL = ""


class SectionCategory(Enum):
    """Enum for section categories."""

    UNKNOWN = -1
    STANDARD = 1
    TAG_SAMPLE = 2
    COMPARISON = 3


@dataclass
class TableSection:
    """Contains information for writing and formatting a section of a table.

    Note that the `data` DataFrame must not contain any NaN values.
    """

    data: pd.DataFrame
    column_formats: dict = field(default_factory=dict)
    column_conditionals: dict = field(default_factory=dict)
    column_widths: dict = field(default_factory=dict)
    headers: dict = field(default_factory=dict)
    header_formats: dict = field(default_factory=dict)
    supheader: str = ""
    supheader_format: dict = field(default_factory=dict)
    section_conditional: str = ""

    def __post_init__(self):
        nan_columns = self.data.columns[self.data.isnull().any()].tolist()
        if nan_columns:
            raise ValueError(f"`data` contains NaN values in columns: {nan_columns}")

        for col in self.data.columns:
            if col not in self.column_formats:
                self.column_formats[col] = {}
            if col not in self.column_conditionals:
                self.column_conditionals[col] = {}
            if col not in self.column_widths:
                self.column_widths[col] = DEFAULT_COL_WIDTH
            if col not in self.headers:
                self.headers[col] = col
            if col not in self.header_formats:
                self.header_formats[col] = {}


class SectionCompiler(Protocol):
    """Protocol for section compilers."""

    def compile(self, section_template: dict, table: pd.DataFrame) -> TableSection:
        """Compile a table section from a section template and a table."""


class StandardSectionCompiler:
    """Compiler for standard table sections."""

    def __init__(self, report_template: ReportTemplate):
        self.formats = report_template.formats
        self.conditional_formats = report_template.conditional_formats
        self.settings = report_template.settings

    def compile(self, section_template: dict, table: pd.DataFrame) -> TableSection:
        """Compile a table section from a standard section template and a table."""
        selected_cols = eval_standard_section_columns(section_template, table.columns)
        data = eval_data(table, selected_cols, section_template)
        col_formats = eval_column_formats(
            selected_cols, section_template, self.formats, DEFAULT_FORMAT
        )
        col_conditionals = eval_column_conditional_formats(
            selected_cols, section_template, self.conditional_formats
        )
        default_width = self.settings.get("column_width", DEFAULT_COL_WIDTH)
        col_widths = eval_column_widths(selected_cols, section_template, default_width)
        headers = {c: c for c in selected_cols}
        header_formats = eval_header_formats(
            selected_cols, section_template, self.formats
        )
        supheader = section_template.get("supheader", "")
        supheader_format = eval_supheader_format(section_template, self.formats)
        section_conditional = eval_section_conditional_format(
            section_template, self.conditional_formats
        )

        return TableSection(
            data=data,
            column_formats=col_formats,
            column_conditionals=col_conditionals,
            column_widths=col_widths,
            headers=headers,
            header_formats=header_formats,
            supheader=supheader,
            supheader_format=supheader_format,
            section_conditional=section_conditional,
        )


class TagSampleSectionCompiler:
    """Compiler for tag sample table sections."""

    def __init__(self, report_template: ReportTemplate):
        self.formats = report_template.formats
        self.conditional_formats = report_template.conditional_formats
        self.settings = report_template.settings

    def compile(self, section_template: dict, table: pd.DataFrame) -> TableSection:
        """Compile a table section from a standard section template and a table."""
        selected_cols = eval_tag_sample_section_columns(
            section_template, table.columns, self.settings["sample_extraction_tag"]
        )
        data = eval_data(table, selected_cols, section_template)
        col_formats = eval_column_formats(
            selected_cols, section_template, self.formats, DEFAULT_FORMAT
        )
        col_conditionals = eval_column_conditional_formats(
            selected_cols, section_template, self.conditional_formats
        )
        default_width = self.settings.get("column_width", DEFAULT_COL_WIDTH)
        col_widths = eval_column_widths(selected_cols, section_template, default_width)
        headers = eval_tag_sample_headers(
            selected_cols, section_template, self.settings.get("log2_tag", "")
        )
        header_formats = eval_header_formats(
            selected_cols, section_template, self.formats
        )
        supheader = eval_tag_sample_supheader(
            section_template, self.settings.get("log2_tag", "")
        )
        supheader_format = eval_supheader_format(section_template, self.formats)
        section_conditional = eval_section_conditional_format(
            section_template, self.conditional_formats
        )

        return TableSection(
            data=data,
            column_formats=col_formats,
            column_conditionals=col_conditionals,
            column_widths=col_widths,
            headers=headers,
            header_formats=header_formats,
            supheader=supheader,
            supheader_format=supheader_format,
            section_conditional=section_conditional,
        )


def get_section_compiler(section_template: dict) -> SectionCompiler:
    """Get the section compiler function for a section template."""
    section_category = identify_template_section_category(section_template)
    if section_category == SectionCategory.UNKNOWN:
        raise ValueError("Unknown section category.")
    elif section_category == SectionCategory.STANDARD:
        return StandardSectionCompiler
    elif section_category == SectionCategory.TAG_SAMPLE:
        return TagSampleSectionCompiler
    elif section_category == SectionCategory.COMPARISON:
        raise NotImplementedError(f"Compiler not implemented for {section_category}.")


# Missing from the compile_table_sections:
# 1) Apply section type specific data manipulations (e.g. log2 transformation)
# 2) Apply common data manipulations (e.g. replace missing values / NaNs)


def prepare_table_sections(
    report_template: ReportTemplate,
    table: pd.DataFrame,
    remove_duplicate_columns: bool = True,
) -> list[TableSection]:
    """Compile non-empty table sections from a report template and a table.

    Args:
        report_template: The report template describing how table sections should be
            generated.
        table: The table to compile the sections from.
        remove_duplicate_columns: If True, duplicate columns are removed from the
            sections, keeping only the first occurrence of a column.

    Returns:
        A list of non-empty, compiled table sections.
    """
    compiled_table_sections = compile_table_sections(report_template, table)
    if remove_duplicate_columns:
        prune_table_sections(compiled_table_sections)
    return remove_empty_table_sections(compiled_table_sections)


def compile_table_sections(
    report_template: ReportTemplate, table: pd.DataFrame
) -> list[TableSection]:
    """Compile table sections from a report template and a table.

    Args:
        report_template: The report template describing how table sections should be
            generated.
        table: The table to compile the sections from.

    Returns:
        A list of compiled table sections.
    """
    table_sections = []
    for section_template in report_template.sections.values():
        section_category = identify_template_section_category(section_template)
        if section_category == SectionCategory.UNKNOWN:
            continue

        SectionCompilerClass = get_section_compiler(section_template)
        section_compiler = SectionCompilerClass(report_template)
        compiled_section = section_compiler.compile(section_template, table)
        if isinstance(compiled_section, TableSection):
            table_sections.append(compiled_section)
        else:
            table_sections.extend(compiled_section)

    return table_sections


def prune_table_sections(table_sections: Iterable[TableSection]) -> None:
    """Remove duplicate columns from table sections, keeping only the first occurance."""
    observed_columns = set()
    for section in table_sections:
        to_remove = [col for col in section.data.columns if col in observed_columns]
        section.data = section.data.drop(columns=to_remove)
        for col in to_remove:
            del section.column_formats[col]
            del section.column_conditionals[col]
            del section.column_widths[col]
            del section.headers[col]
            del section.header_formats[col]
        observed_columns.update(section.data.columns)


def remove_empty_table_sections(
    table_sections: Iterable[TableSection],
) -> list[TableSection]:
    """Returns a list of non-empty table sections."""
    return [section for section in table_sections if not section.data.empty]


def eval_data(table: pd.DataFrame, columns: Iterable[str], section_template: dict):
    """Returns a copy of the table with only the selected columns and no NaN values.

    Args:
        table: The table to select columns from.
        columns: The columns to select from the table.
        section_template:
    """
    data = table[columns].copy()
    if section_template.get("log2", False):
        if not data.select_dtypes(exclude=["number"]).columns.empty:
            raise ValueError("Cannot log2 transform non-numeric columns.")
        data = data.mask(data <= 0, np.nan)
        data = np.log2(data)
    data.fillna(NAN_REPLACEMENT_SYMBOL, inplace=True)
    return data


def eval_standard_section_columns(
    section_template: dict, columns: Iterable[str]
) -> list[str]:
    """Select columns from the template that are present in the table.

    Args:
        section_template: A dictionary containing the columns to be selected as the
            values of the "columns" key.
        columns: A list of column names to select from.

    Returns:
        A list of column names that are present in both the template and the table.
    """
    selected_columns = [col for col in section_template["columns"] if col in columns]
    return selected_columns


def eval_tag_sample_section_columns(
    section_template: dict, columns: Iterable[str], extraction_tag: str
) -> list[str]:
    """Extract tag sample columns.

    Args:
        section_template: A dictionary containing the columns to be selected as the
            values of the "columns" key.
        columns: A list of column names to select from.
        extraction_tag: The tag used to extract sample names from the columns.

    Returns:
        A list of sample columns that contain the `section_template["tag"]`.
    """
    samples = []
    for col in columns:
        if extraction_tag not in col or col == extraction_tag:
            continue
        samples.append(col.replace(extraction_tag, "").strip())

    selected_columns = []
    for col in columns:
        if section_template["tag"] not in col:
            continue
        for sample in samples:
            if sample in col:
                selected_columns.append(col)
    return selected_columns


def eval_tag_sample_headers(
    columns: Iterable[str],
    section_template: dict,
    log2_tag: str = "",
) -> dict:
    """Returns header names for each column.

    Args:
        columns: A list of column names to select from.
        section_template: A dictionary with "tag" containing the substring that will be
            removed from the headers if "remove_tag" is True. The "log2" key determines
            whether to add the `log2_tag` to the headers, however, if "remove_tag" is
            True the `log2_tag` will never be added. The "remove_tag" and "log"
            keys are optional and by default False.
        log2_tag: The substring that will be added to the column names if `log2` is
            True.

    Returns:
        A dictionary containing the header names for each column.
    """
    tag = section_template["tag"]
    remove_tag = section_template.get("remove_tag", False)
    add_log2_tag = section_template.get("log2", False) and not remove_tag
    if remove_tag:
        headers = {col: col.replace(tag, "").strip() for col in columns}
    else:
        headers = {col: col for col in columns}

    if add_log2_tag:
        headers = {c: f"{h} {log2_tag}" for c, h in headers.items()}
    return headers


def eval_tag_sample_supheader(
    section_template: dict,
    log2_tag: str = "",
) -> str:
    """Returns header names for each column.

    Args:
        columns: A list of column names to select from.
        section_template: A dictionary with "tag" containing the substring that will be
            removed from the headers if "remove_tag" is True. The "log2" key determines
            whether to add the `log2_tag` to the headers, however, if "remove_tag" is
            True the `log2_tag` will never be added. The "remove_tag" and "log"
            keys are optional and by default False.
        log2_tag: The substring that will be added to the column names if `log2` is
            True.

    Returns:
        A dictionary containing the header names for each column.
    """
    if section_template.get("log2", False):
        return f"{section_template['supheader']} {log2_tag}"
    else:
        return section_template["supheader"]


def eval_column_formats(
    columns: str,
    section_template: dict,
    format_templates: dict,
    default_format: Optional[dict] = None,
) -> dict:
    """Returns format descriptions for each column in the section.

    If "border" is set to True in the `section_template`, the format descriptions for
    the first and last column are updated to include borders.

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
    if section_template.get("border", False):
        column_formats[columns[0]]["left"] = BORDER_TYPE
        column_formats[columns[-1]]["right"] = BORDER_TYPE
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
    default_width: float = 64,
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
    `format_templates`. If "border" is set to True in the `section_template`, the
    header format descriptions for the first and last column are updated to include
    borders.

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
    if section_template.get("border", False):
        column_header_formats[columns[0]]["left"] = BORDER_TYPE
        column_header_formats[columns[-1]]["right"] = BORDER_TYPE
    return column_header_formats


def eval_supheader_format(section_template: dict, format_templates: dict) -> dict:
    """Returns a format descriptions for the supheader.

    Supheader format description defined in the `section_template` updates the one from
    the `format_templates`. If "border" is set to True in the `section_template`, the
    supheader format description is updated to include borders.

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
    if section_template.get("border", False):
        supheader_format.update({"left": BORDER_TYPE, "right": BORDER_TYPE})

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


def identify_template_section_category(section_template: dict) -> SectionCategory:
    """Identify the category of a section template.

    Args:
        section_template: A dictionary containing the section template.

    Returns:
        A SectionCategory enum value.
    """
    has_comp_group = "comparison_group" in section_template
    has_tag = "tag" in section_template
    has_columns = "columns" in section_template

    if has_comp_group and not has_columns:
        return SectionCategory.COMPARISON
    if has_tag and not has_comp_group and not has_columns:
        return SectionCategory.TAG_SAMPLE
    if has_columns and not has_tag and not has_comp_group:
        return SectionCategory.STANDARD
    return SectionCategory.UNKNOWN
