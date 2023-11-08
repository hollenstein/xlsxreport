"""Contains functions for compiling table sections from a report template and a table."""
from dataclasses import dataclass, field
from enum import Enum
from typing import Iterable, Optional, Protocol

import numpy as np
import pandas as pd

from xlsxreport.template import ReportTemplate


BORDER_TYPE: int = 2  # 2 = thick line, see xlsxwriter.format.Format().set_border()
DEFAULT_COL_WIDTH: float = 64
DEFAULT_FORMAT: dict = {"num_format": "@"}
REMAINING_COL_FORMAT = {"align": "left", "num_format": "0"}
NAN_REPLACEMENT_SYMBOL = ""
WHITESPACE_CHARS = " ."


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
        selected_cols = eval_standard_section_columns(table.columns, section_template)
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
            table.columns, section_template, self.settings["sample_extraction_tag"]
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


class ComparisonSectionCompiler:
    """Compiler for comparison table sections."""

    def __init__(self, report_template: ReportTemplate):
        self.standard_compiler = StandardSectionCompiler(report_template)

    def compile(
        self, section_template: dict, table: pd.DataFrame
    ) -> list[TableSection]:
        """Compile table sections from a comparison section template and a table."""

        comparison_groups = eval_comparison_groups(table.columns, section_template)
        table_sections = []
        for comparison_group in comparison_groups:
            selected_cols = eval_comparison_group_columns(
                table.columns, section_template, comparison_group
            )
            col_conditionals = eval_comparison_group_conditional_format_names(
                selected_cols, section_template
            )
            supheader = eval_comparison_group_supheader(
                section_template, comparison_group
            )

            std_section_template = section_template.copy()
            std_section_template["columns"] = selected_cols
            std_section_template["column_conditional"] = col_conditionals
            std_section_template["supheader"] = supheader

            table_section = self.standard_compiler.compile(std_section_template, table)
            table_section.headers = eval_comparison_group_headers(
                selected_cols, section_template, comparison_group
            )
            table_sections.append(table_section)

        return table_sections


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
        return ComparisonSectionCompiler


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
    if report_template.settings.get("append_remaining_columns", False):
        remaining_section = compile_remaining_column_table_section(
            report_template, compiled_table_sections, table
        )
        compiled_table_sections.append(remaining_section)
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


def compile_remaining_column_table_section(
    report_template: ReportTemplate,
    table_sections: Iterable[TableSection],
    table: pd.DataFrame,
) -> TableSection:
    """Compile a table section containing all columns not present in other sections.

    Args:
        report_template: The report template describing how table sections should be
            generated.
        table_sections: The table sections that have already been compiled.
        table: The table to compile the remaining column section from.

    Returns:
        A compiled table section containing all columns not present in other sections.
    """
    observed_columns = set()
    for section in table_sections:
        observed_columns.update(section.data.columns)
    selected_cols = [column for column in table if column not in observed_columns]

    section_compiler = StandardSectionCompiler(report_template)
    section_compiler.formats["__remaining__"] = REMAINING_COL_FORMAT
    section_template = {
        "columns": selected_cols,
        "format": "__remaining__",
        "width": DEFAULT_COL_WIDTH,
    }
    section = section_compiler.compile(section_template, table)
    return section


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
    columns: Iterable[str], section_template: dict
) -> list[str]:
    """Select columns from the template that are present in the table.

    Args:
        columns: A list of column names to select from.
        section_template: A dictionary containing the columns to be selected as the
            values of the "columns" key.

    Returns:
        A list of column names that are present in both the template and the table.
    """
    selected_columns = [col for col in section_template["columns"] if col in columns]
    return selected_columns


def eval_tag_sample_section_columns(
    columns: Iterable[str], section_template: dict, extraction_tag: str
) -> list[str]:
    """Extract tag sample columns.

    Args:
        columns: A list of column names to select from.
        section_template: A dictionary containing the columns to be selected as the
            values of the "columns" key.
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


def eval_comparison_groups(columns: Iterable[str], section_template: dict) -> list[str]:
    """Extract comparison groups from the columns of a table.

    Args:
        columns: A list of column names used for extracting comparison group names.
        section_template: A dictionary containing the keys "tag" (a string) and
            "columns" (a list of strings).

    Returns:
        A list of unique comparison group names. Comparison group names are extracted
        from the `columns` that contain the substring specified in the
        `section_template` by "tag" and one of the substrings specified by "columns".
        Comparison groups are extracted by removing the `section_template["columns"]`
        substrings from the column name and stripping whitespace characters from the
        result. Each comparison group name is extracted only once.
    """
    comparison_tag = section_template["tag"]
    comparison_columns = [col for col in columns if comparison_tag in col]

    comparison_groups = []
    for column_tag in section_template["columns"]:
        matched_columns = [col for col in comparison_columns if column_tag in col]
        for column in matched_columns:
            putative_group = column.replace(column_tag, "").strip(WHITESPACE_CHARS)
            if putative_group and putative_group not in comparison_groups:
                comparison_groups.append(putative_group)
    return comparison_groups


def eval_comparison_group_columns(
    columns: Iterable[str], section_template: dict, comparison_group: str
) -> list[str]:
    """Extract columns from a table that belong to a comparison group.

    Args:
        columns: A list of column names to select from.
        section_template: A dictionary containing the key "columns" with its value
            being a list of strings.
        comparison_group: The name of the comparison group to extract columns for.

    Returns:
        A list of column names that consist of the `comparison_group` and one of the
        substrings specified by `section_template["columns"]`.
    """
    selected_columns = []
    for column in [col for col in columns if comparison_group in col]:
        leftover = column.replace(comparison_group, "")
        for column_tag in section_template["columns"]:
            leftover = leftover.replace(column_tag, "")
        if leftover.strip(WHITESPACE_CHARS) == "":
            selected_columns.append(column)
    return selected_columns


def eval_comparison_group_headers(
    columns: Iterable[str], section_template: dict, comparison_group: str
):
    """Returns header names for each column.

    Args:
        columns: A list of column names to select from.
        section_template: A dictionary with keys "remove_tag" (bool), "tag" (str), and
            "replace_comparison_tag" (str). If "remove_tag" is True, the
            `comparison_group` will be removed from the headers. If the
            "replace_comparison_tag" is specified, the "tag" value will be replaced with
            the value of "replace_comparison_tag" in the headers.
        comparison_group: The name of the comparison group.

    Returns:
        A dictionary containing the header names for each column.
    """
    remove_comparison_group = section_template.get("remove_tag", False)
    old_comparison_tag = section_template["tag"]
    new_comparison_tag = section_template.get("replace_comparison_tag", None)
    replace_comparison_tag = True if new_comparison_tag is not None else False

    headers = {}
    for col in columns:
        header = col
        if remove_comparison_group:
            header = header.replace(comparison_group, "").strip(WHITESPACE_CHARS)
        if replace_comparison_tag:
            header = header.replace(old_comparison_tag, new_comparison_tag)
        headers[col] = header
    return headers


def eval_comparison_group_supheader(section_template: dict, comparison_group: str):
    """Returns the supheader name for a comparison group.

    Args:
        section_template: A dictionary with keys "tag" (str) and
            "replace_comparison_tag" (str). If the "replace_comparison_tag" is
            specified, the "tag" value will be replaced with the value of
            "replace_comparison_tag" in the supheader.
        comparison_group: The name used for the supheader.

    Returns:
        The supheader name for the comparison group section.
    """
    old_comparison_tag = section_template["tag"]
    new_comparison_tag = section_template.get("replace_comparison_tag", None)
    replace_comparison_tag = True if new_comparison_tag is not None else False

    supheader = comparison_group
    if replace_comparison_tag:
        supheader = supheader.replace(old_comparison_tag, new_comparison_tag)
    return supheader


def eval_comparison_group_conditional_format_names(
    columns: Iterable[str], section_template: dict
) -> dict[str, str]:
    """Returns conditional format names for each column.

    Conditional formats are matched to each column by checking if any of the
    `section_template["conditional format"]` keys are a substring of the column name,
    and using the corresponding value as the conditional format name for the column.

    Args:
        columns: A list of column names.
        section_template: A dictionary with a "column_conditional" key containing a
            dictionary with column names as keys and conditional format names as values.

    Returns:
        A dictionary containing conditional format names for each column.
    """
    conditional_formats = section_template.get("column_conditional", {})
    col_conditionals = {}
    for tag, format_name in conditional_formats.items():
        for column in columns:
            if tag in column:
                col_conditionals[column] = format_name
    return col_conditionals


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
    """Returns the supheader name for a tag sample section.

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
        The supheader name for the section.
    """
    supheader = section_template.get("supheader", section_template["tag"])
    if section_template.get("log2", False):
        supheader = f"{supheader} {log2_tag}"
    return supheader


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
    if not columns:
        return {}
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
    if not columns:
        return {}
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
    is_comp_group = section_template.get("comparison_group", False)
    has_tag = "tag" in section_template
    has_columns = "columns" in section_template

    if is_comp_group and has_columns and has_tag:
        return SectionCategory.COMPARISON
    if has_tag and not has_columns and not is_comp_group:
        return SectionCategory.TAG_SAMPLE
    if has_columns and not has_tag and not is_comp_group:
        return SectionCategory.STANDARD
    return SectionCategory.UNKNOWN
