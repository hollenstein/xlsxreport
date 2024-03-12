"""Contains functions for compiling table sections from a table template and a table."""

from __future__ import annotations
from copy import deepcopy
from dataclasses import dataclass, field
from typing import Iterable, Optional, Protocol, Sequence
from collections.abc import Mapping, MutableMapping
import re

import numpy as np
import pandas as pd

from xlsxreport.template import SectionCategory


BORDER_TYPE: int = 2  # 2 = thick line, see xlsxwriter.format.Format().set_border()
DEFAULT_COL_WIDTH: float = 64
DEFAULT_FORMAT = {"num_format": "General"}
REMAINING_COL_FORMAT = {"num_format": "General"}
NAN_REPLACEMENT_SYMBOL = ""
WHITESPACE_CHARS = " ."


class TableTemplate(Protocol):
    """Abstract class representing a table template."""

    @property
    def sections(self) -> Mapping[str, TemplateSection]: ...

    @property
    def formats(self) -> MutableMapping[str, Mapping]: ...

    @property
    def conditional_formats(self) -> Mapping[str, Mapping]: ...

    @property
    def settings(self) -> Mapping: ...


class TemplateSection(Protocol):
    """Abstract class representing a section of a table template."""

    category: SectionCategory

    def to_dict(self) -> dict: ...


class SectionCompiler(Protocol):
    """Protocol for section compilers."""

    def __init__(self, table_template: TableTemplate): ...

    def compile(
        self, section_template: dict, table: pd.DataFrame
    ) -> list[CompiledSection]:
        """Compile a table section from a section template and a table."""
        ...


@dataclass
class CompiledSection:
    """Contains information for writing and formatting a section of a table.

    Note that the `data` DataFrame must not contain any NaN values.
    """

    data: pd.DataFrame
    column_formats: dict = field(default_factory=dict)
    column_conditional_formats: dict = field(default_factory=dict)
    column_widths: dict = field(default_factory=dict)
    headers: dict = field(default_factory=dict)
    header_formats: dict = field(default_factory=dict)
    supheader: str = ""
    supheader_format: dict = field(default_factory=dict)
    section_conditional_format: dict = field(default_factory=dict)
    hide_section: bool = False

    def __post_init__(self):
        nan_columns = self.data.columns[self.data.isnull().any()].tolist()
        if nan_columns:
            raise ValueError(
                f"Compiled section contains NaN values in columns: {nan_columns}"
            )
        if self.data.columns.size != self.data.columns.nunique():
            duplicates = self.data.columns[self.data.columns.duplicated()].unique()
            duplicate_message = ", ".join([f"'{c}'" for c in duplicates])
            raise ValueError(
                f"Compiled section contains duplicate columns: {duplicate_message}"
            )

        for col in self.data.columns:
            if col not in self.column_formats:
                self.column_formats[col] = {}
            if col not in self.column_conditional_formats:
                self.column_conditional_formats[col] = {}
            if col not in self.column_widths:
                self.column_widths[col] = DEFAULT_COL_WIDTH
            if col not in self.headers:
                self.headers[col] = col
            if col not in self.header_formats:
                self.header_formats[col] = {}


class StandardSectionCompiler:
    """Compiler for standard table sections."""

    def __init__(self, table_template: TableTemplate):
        self.formats = table_template.formats
        self.conditional_formats = table_template.conditional_formats
        self.settings = table_template.settings

    def compile(
        self, section_template: Mapping, table: pd.DataFrame
    ) -> list[CompiledSection]:
        """Compile a table section from a standard section template and a table."""
        selected_cols = eval_standard_section_columns(table.columns, section_template)
        data = eval_data(table, selected_cols)
        col_formats = eval_column_formats(
            selected_cols, section_template, self.formats, DEFAULT_FORMAT
        )
        col_conditionals = eval_column_conditional_formats(
            selected_cols, section_template, self.conditional_formats
        )
        default_width = self.settings["column_width"]
        col_widths = eval_column_widths(selected_cols, section_template, default_width)
        headers = {c: c for c in selected_cols}
        header_formats = eval_header_formats(
            selected_cols, section_template, self.formats
        )
        supheader = section_template.get("supheader", "")
        supheader_format = eval_supheader_format(section_template, self.formats)
        section_conditional_format = eval_section_conditional_format(
            section_template, self.conditional_formats
        )
        hide_section = section_template.get("hide_section", False)
        compiled_section = CompiledSection(
            data=data,
            column_formats=col_formats,
            column_conditional_formats=col_conditionals,
            column_widths=col_widths,
            headers=headers,
            header_formats=header_formats,
            supheader=supheader,
            supheader_format=supheader_format,
            section_conditional_format=section_conditional_format,
            hide_section=hide_section,
        )
        return [compiled_section]


class TagSectionCompiler:
    """Compiler for tag table sections."""

    def __init__(self, table_template: TableTemplate):
        self.formats = table_template.formats
        self.conditional_formats = table_template.conditional_formats
        self.settings = table_template.settings

    def compile(
        self, section_template: Mapping, table: pd.DataFrame
    ) -> list[CompiledSection]:
        """Compile a table section from a standard section template and a table."""
        selected_cols = eval_tag_section_columns(table.columns, section_template)
        data = eval_data_with_log2_transformation(
            table,
            selected_cols,
            section_template,
            self.settings["evaluate_log2_transformation"],
        )
        col_formats = eval_column_formats(
            selected_cols, section_template, self.formats, DEFAULT_FORMAT
        )
        col_conditionals = eval_column_conditional_formats(
            selected_cols, section_template, self.conditional_formats
        )
        default_width = self.settings["column_width"]
        col_widths = eval_column_widths(selected_cols, section_template, default_width)
        headers = eval_tag_sample_headers(
            selected_cols, section_template, self.settings["log2_tag"]
        )
        header_formats = eval_header_formats(
            selected_cols, section_template, self.formats
        )
        supheader = eval_tag_sample_supheader(
            section_template, self.settings["log2_tag"]
        )
        supheader_format = eval_supheader_format(section_template, self.formats)
        section_conditional_format = eval_section_conditional_format(
            section_template, self.conditional_formats
        )
        hide_section = section_template.get("hide_section", False)
        compiled_section = CompiledSection(
            data=data,
            column_formats=col_formats,
            column_conditional_formats=col_conditionals,
            column_widths=col_widths,
            headers=headers,
            header_formats=header_formats,
            supheader=supheader,
            supheader_format=supheader_format,
            section_conditional_format=section_conditional_format,
            hide_section=hide_section,
        )
        return [compiled_section]


class LabelTagSectionCompiler:
    """Compiler for tag table sections."""

    def __init__(self, table_template: TableTemplate):
        self.formats = table_template.formats
        self.conditional_formats = table_template.conditional_formats
        self.settings = table_template.settings

    def compile(
        self, section_template: Mapping, table: pd.DataFrame
    ) -> list[CompiledSection]:
        """Compile a table section from a standard section template and a table."""
        selected_cols = eval_label_tag_section_columns(table.columns, section_template)
        data = eval_data_with_log2_transformation(
            table,
            selected_cols,
            section_template,
            self.settings["evaluate_log2_transformation"],
        )
        col_formats = eval_column_formats(
            selected_cols, section_template, self.formats, DEFAULT_FORMAT
        )
        col_conditionals = eval_column_conditional_formats(
            selected_cols, section_template, self.conditional_formats
        )
        default_width = self.settings["column_width"]
        col_widths = eval_column_widths(selected_cols, section_template, default_width)
        headers = eval_tag_sample_headers(
            selected_cols, section_template, self.settings["log2_tag"]
        )
        header_formats = eval_header_formats(
            selected_cols, section_template, self.formats
        )
        supheader = eval_tag_sample_supheader(
            section_template, self.settings["log2_tag"]
        )
        supheader_format = eval_supheader_format(section_template, self.formats)
        section_conditional_format = eval_section_conditional_format(
            section_template, self.conditional_formats
        )
        hide_section = section_template.get("hide_section", False)
        compiled_section = CompiledSection(
            data=data,
            column_formats=col_formats,
            column_conditional_formats=col_conditionals,
            column_widths=col_widths,
            headers=headers,
            header_formats=header_formats,
            supheader=supheader,
            supheader_format=supheader_format,
            section_conditional_format=section_conditional_format,
            hide_section=hide_section,
        )
        return [compiled_section]


class ComparisonSectionCompiler:
    """Compiler for comparison table sections."""

    def __init__(self, table_template: TableTemplate):
        self.std_compiler = StandardSectionCompiler(table_template)

    def compile(
        self, section_template: Mapping, table: pd.DataFrame
    ) -> list[CompiledSection]:
        """Compile table sections from a comparison section template and a table."""

        comparison_groups = eval_comparison_groups(table.columns, section_template)
        compiled_sections = []
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

            std_section_template = deepcopy(dict(section_template))
            std_section_template["columns"] = selected_cols
            std_section_template["column_conditional_format"] = col_conditionals
            std_section_template["supheader"] = supheader

            compiled_section = self.std_compiler.compile(std_section_template, table)[0]
            compiled_section.headers = eval_comparison_group_headers(
                selected_cols, section_template, comparison_group
            )
            compiled_sections.append(compiled_section)

        return compiled_sections


def get_section_compiler(section_category: SectionCategory) -> type[SectionCompiler]:
    """Get the appropriate section compiler for a section category."""
    if section_category not in _CATEGORY_COMPILER_MAP:
        raise NotImplementedError(
            f"Section compiler not implemented for category {section_category}."
        )
    return _CATEGORY_COMPILER_MAP[section_category]


def prepare_compiled_sections(
    table_template: TableTemplate,
    table: pd.DataFrame,
) -> list[CompiledSection]:
    """Compile non-empty table sections from a table template and a table.

    First the table sections are compiled from the template and the table. If the
    "append_remaining_columns" setting is True, the remaining columns are compiled into
    a section that is appended to the list of compiled table sections. Duplicate columns
    are removed from the table sections if the "remove_duplicate_columns" setting is
    True. Finally, empty table sections are removed from the list of table sections.

    Args:
        table_template: The template describing how table sections should be generated.
        table: The table to compile the sections from.

    Returns:
        A list of non-empty, compiled table sections.
    """
    compiled_sections = compile_sections(table_template, table)
    if table_template.settings["append_remaining_columns"]:
        remaining_section = compile_remaining_column_section(
            table_template, compiled_sections, table
        )
        compiled_sections.append(remaining_section)
    if table_template.settings["remove_duplicate_columns"]:
        prune_compiled_sections(compiled_sections)
    return remove_empty_compiled_sections(compiled_sections)


def compile_sections(
    table_template: TableTemplate, table: pd.DataFrame
) -> list[CompiledSection]:
    """Compile table sections from a table template and a table.

    Args:
        table_template: The template describing how table sections should be generated.
        table: The table to compile the sections from.

    Returns:
        A list of compiled table sections.
    """
    all_compiled_sections = []
    for section in table_template.sections.values():
        section_template = section.to_dict()
        if section.category == SectionCategory.UNKNOWN:
            continue

        _SectionCompiler = get_section_compiler(section.category)
        section_compiler = _SectionCompiler(table_template)
        compiled_sections = section_compiler.compile(section_template, table)
        all_compiled_sections.extend(compiled_sections)

    return all_compiled_sections


def compile_remaining_column_section(
    table_template: TableTemplate,
    compiled_sections: Iterable[CompiledSection],
    table: pd.DataFrame,
) -> CompiledSection:
    """Compile a table section containing all columns not present in other sections.

    Args:
        table_template: The template describing how table sections should be generated.
        compiled_sections: The table sections that have already been compiled.
        table: The table to compile the remaining column section from.

    Returns:
        A compiled table section containing all columns not present in other sections.
    """
    observed_columns: set = set()
    for section in compiled_sections:
        observed_columns.update(section.data.columns)
    selected_cols = [column for column in table if column not in observed_columns]

    section_compiler = StandardSectionCompiler(table_template)
    _format_name = "_" * (max([len(i) for i in section_compiler.formats]) + 1)
    section_compiler.formats[_format_name] = REMAINING_COL_FORMAT
    section_template = {
        "columns": selected_cols,
        "format": _format_name,
        "width": DEFAULT_COL_WIDTH,
        "hide_section": True,
    }
    section = section_compiler.compile(section_template, table)[0]
    del section_compiler.formats[_format_name]
    return section


def prune_compiled_sections(compiled_sections: Iterable[CompiledSection]) -> None:
    """Remove duplicate columns from table sections, keeping only the first occurance."""
    observed_columns: set = set()
    for section in compiled_sections:
        to_remove = [col for col in section.data.columns if col in observed_columns]
        section.data = section.data.drop(columns=to_remove)
        for col in to_remove:
            del section.column_formats[col]
            del section.column_conditional_formats[col]
            del section.column_widths[col]
            del section.headers[col]
            del section.header_formats[col]
        observed_columns.update(section.data.columns)


def remove_empty_compiled_sections(
    compiled_sections: Iterable[CompiledSection],
) -> list[CompiledSection]:
    """Returns a list of non-empty table sections."""
    return [section for section in compiled_sections if not section.data.empty]


def eval_data(table: pd.DataFrame, columns: Iterable[str]) -> pd.DataFrame:
    """Returns a copy of the table with only the selected columns and no NaN values.

    Args:
        table: The table to select columns from.
        columns: The columns to select from the table.

    Returns:
        A copy of the table with only the selected columns NaN values replaced.
    """
    data = table[columns]
    nan_values = data.isna()
    data = data.astype("object")
    data[nan_values] = NAN_REPLACEMENT_SYMBOL
    return data


def eval_data_with_log2_transformation(
    table: pd.DataFrame,
    columns: Iterable[str],
    section_template: Mapping,
    evaluate_log_state: bool,
) -> pd.DataFrame:
    """Selects columns from the table and applies a log2 transformation if specified.

    Args:
        table: The table to select columns from.
        columns: The columns to select from the table.
        section_template: A dictionary containing the key "log2" (bool), which
            determines whether to apply a log2 transformation.
        evaluate_log_state: If True, the log2 transformation is only applied if the
            intensities are not already in log space.

    Returns:
        A copy of the table with only the selected columns and NaN vaues replaced,
        optionally values are log2 transformed.
    """
    data = table[columns].copy()
    apply_log_transform = section_template.get("log2", False)
    if apply_log_transform and evaluate_log_state and _intensities_in_logspace(data):
        apply_log_transform = False

    if apply_log_transform:
        if not data.select_dtypes(exclude=["number"]).columns.empty:
            raise ValueError("Cannot log2 transform non-numeric columns.")
        data = data.mask(data <= 0, np.nan)
        data = np.log2(data)  # type: ignore

    nan_values = data.isna()
    data = data.astype("object")
    data[nan_values] = NAN_REPLACEMENT_SYMBOL
    return data


def eval_standard_section_columns(
    columns: Iterable[str], section_template: Mapping
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


def eval_tag_section_columns(
    columns: Iterable[str], section_template: Mapping
) -> list[str]:
    """Extract columns using a regex pattern.

    Args:
        columns: A list of column names to select from.
        section_template: A dictionary containing the key "tag", which is a regular
            expression pattern used to select the section columns.

    Returns:
        A list of sample columns matching the pattern in `section_template["tag"]`.
    """
    selected_columns = [c for c in columns if re.search(section_template["tag"], c)]
    return selected_columns


def eval_label_tag_section_columns(
    columns: Iterable[str], section_template: Mapping
) -> list[str]:
    """Extract columns using a regex pattern.

    Args:
        columns: A list of column names to select from.
        section_template: A dictionary containing the key "tag", which is a regular
            expression pattern used to select the section columns.

    Returns:
        A list of sample columns matching the pattern in `section_template["tag"]`.
    """
    selected_columns = []
    column_query = [c for c in columns if re.search(section_template["tag"], c)]
    for column in column_query:
        match = re.sub(section_template["tag"], "", column).strip(WHITESPACE_CHARS)
        if match in section_template["labels"]:
            selected_columns.append(column)
    return selected_columns


def eval_comparison_groups(
    columns: Iterable[str], section_template: Mapping
) -> list[str]:
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
    columns: Iterable[str], section_template: Mapping, comparison_group: str
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
    selected_cols = []
    comparison_group_cols = [col for col in columns if comparison_group in col]
    for column_tag in section_template["columns"]:
        for column in comparison_group_cols:
            leftover = column.replace(comparison_group, "").replace(column_tag, "")
            if leftover.strip(WHITESPACE_CHARS) == "" and column not in selected_cols:
                selected_cols.append(column)
                break
    return selected_cols


def eval_comparison_group_headers(
    columns: Iterable[str], section_template: Mapping, comparison_group: str
) -> dict[str, str]:
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


def eval_comparison_group_supheader(
    section_template: Mapping, comparison_group: str
) -> str:
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
    columns: Iterable[str], section_template: Mapping
) -> dict[str, str]:
    """Returns conditional format names for each column.

    Conditional formats are matched to each column by checking if any of the
    `section_template["conditional format"]` keys are a substring of the column name,
    and using the corresponding value as the conditional format name for the column.

    Args:
        columns: A list of column names.
        section_template: A dictionary with a "column_conditional_format" key containing
            a dictionary with column names as keys and conditional format names as
            values.

    Returns:
        A dictionary containing conditional format names for each column.
    """
    conditional_formats = section_template.get("column_conditional_format", {})
    col_conditionals = {}
    for tag, format_name in conditional_formats.items():
        for column in columns:
            if tag in column:
                col_conditionals[column] = format_name
    return col_conditionals


def eval_tag_sample_headers(
    columns: Iterable[str],
    section_template: Mapping,
    log2_tag: str = "",
) -> dict[str, str]:
    """Returns header names for each column.

    Args:
        columns: A list of column names to select from.
        section_template: A dictionary with "tag" containing the regex pattern that will
            be removed from the headers if "remove_tag" is True. The "log2" key
            determines whether to add the `log2_tag` to the headers, however, if
            "remove_tag" is True the `log2_tag` will never be added. The "remove_tag"
            and "log" keys are optional and by default False.
        log2_tag: The substring that will be added to the column names if `log2` is
            True.

    Returns:
        A dictionary containing the header names for each column.
    """
    tag = section_template["tag"]
    remove_tag = section_template.get("remove_tag", False)
    add_log2_tag = section_template.get("log2", False) and not remove_tag and log2_tag
    if remove_tag:
        headers = {col: re.sub(tag, "", col).strip() for col in columns}
    else:
        headers = {col: col for col in columns}

    if add_log2_tag and log2_tag:
        headers = {c: f"{h} {log2_tag}" for c, h in headers.items()}
    return headers


def eval_tag_sample_supheader(
    section_template: Mapping,
    log2_tag: str,
) -> str:
    """Returns the supheader name for a tag sample section.

    Args:
        section_template: A dictionary with "supheader" containing the supheader name,
            if "supheader" is not present an empty string is used. The "log2" key
            determines whether to add the `log2_tag` to the supheader name. The "log"
            key is optional and by default False.
        log2_tag: The substring that will be added to the column names if
            `section_template["log"]` is True.

    Returns:
        The supheader name for the section.
    """
    add_log2_tag = section_template.get("log2", False) and log2_tag
    supheader = section_template.get("supheader", "")
    if add_log2_tag and supheader:
        supheader = f"{supheader} {log2_tag}"
    return supheader


def eval_column_formats(
    columns: Sequence[str],
    section_template: Mapping,
    format_templates: Mapping[str, Mapping],
    default_format: Optional[dict] = None,
) -> dict[str, dict]:
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

        column_formats[col] = dict(format_templates.get(format_name, default_format))
    if section_template.get("border", False):
        column_formats[columns[0]]["left"] = BORDER_TYPE
        column_formats[columns[-1]]["right"] = BORDER_TYPE
    return column_formats


def eval_column_conditional_formats(
    columns: Iterable[str],
    section_template: Mapping,
    format_templates: Mapping[str, Mapping],
) -> dict[str, dict]:
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
    column_formats: dict = {}
    for col in columns:
        col_format: dict = {}
        if "column_conditional_format" in section_template:
            if col in section_template["column_conditional_format"]:
                format_name = section_template["column_conditional_format"][col]
                col_format = dict(format_templates.get(format_name, col_format))
        column_formats[col] = col_format
    return column_formats


def eval_column_widths(
    columns: Iterable[str],
    section_template: Mapping,
    default_width: float = 64,
) -> dict[str, float]:
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
    columns: Sequence[str],
    section_template: Mapping,
    format_templates: Mapping[str, Mapping],
) -> dict[str, dict]:
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
    template_format = format_templates.get("header", {})
    section_format = section_template.get("header_format", {})
    header_format = dict(template_format, **section_format)
    column_header_formats = {col: header_format.copy() for col in columns}
    if section_template.get("border", False):
        column_header_formats[columns[0]]["left"] = BORDER_TYPE
        column_header_formats[columns[-1]]["right"] = BORDER_TYPE
    return column_header_formats


def eval_supheader_format(
    section_template: Mapping, format_templates: Mapping[str, Mapping]
) -> dict:
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
    template_format = format_templates.get("supheader", {})
    section_format = section_template.get("supheader_format", {})
    supheader_format = dict(template_format, **section_format)
    if section_template.get("border", False):
        supheader_format.update({"left": BORDER_TYPE, "right": BORDER_TYPE})

    return supheader_format


def eval_section_conditional_format(
    section_template: Mapping, format_templates: Mapping[str, Mapping]
) -> dict:
    """Returns a conditional format description of a section.

    Args:
        section_template: A dictionary that can contain a conditional format name with
            the key "conditional_format".
        format_templates: A dictionary containing the conditional format descriptions
            for each conditional format name. If a format name is not present in the
            `format_templates`, an empty dictionary is used instead.

    Returns:
        A dictionary containing a conditional format description.
    """
    section_format_name = section_template.get("conditional_format", None)
    section_conditional_format = dict(format_templates.get(section_format_name, {}))
    return section_conditional_format


def _intensities_in_logspace(data: pd.DataFrame | np.ndarray | Iterable) -> np.bool_:
    """Evaluates whether intensities are likely to be log transformed.

    Assumes that intensities are log transformed if all values are smaller or
    equal to 64. Intensities values (and intensity peak areas) reported by
    tandem mass spectrometery typically range from 10^3 to 10^12. To reach log2
    transformed values greather than 64, intensities would need to be higher
    than 10^19, which seems to be very unlikely to be ever encountered.

    Args:
        data: Dataset that contains only intensity values, can be any iterable,
            a numpy.array or a pandas.DataFrame, multiple dimensions or columns
            are allowed.

    Returns:
        True if intensity values in 'data' appear to be log transformed.
    """
    data = np.array(data, dtype=float)
    mask = np.isfinite(data)
    return np.all(data[mask].flatten() <= 64)


_CATEGORY_COMPILER_MAP: dict[SectionCategory, type[SectionCompiler]] = {
    SectionCategory.STANDARD: StandardSectionCompiler,
    SectionCategory.TAG: TagSectionCompiler,
    SectionCategory.LABEL_TAG: LabelTagSectionCompiler,
    SectionCategory.COMPARISON: ComparisonSectionCompiler,
}
