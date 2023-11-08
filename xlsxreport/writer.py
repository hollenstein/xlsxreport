from __future__ import annotations
import dataclasses
from typing import Iterable, Optional, Union

import numpy as np
import pandas as pd
import xlsxwriter

from xlsxreport.template import ReportTemplate


WHITESPACE_SYMBOLS = " ._-"


class Reportbook(xlsxwriter.Workbook):
    """Subclass of the XlsxWriter Workbook class."""

    def add_infosheet(self) -> xlsxwriter.worksheet.Worksheet:
        worksheet = self.add_worksheet("Info")
        return worksheet

    def add_datasheet(self, name: Optional[str] = None) -> Datasheet:
        worksheet = self.add_worksheet(name)
        data_sheet = Datasheet(self, worksheet)
        return data_sheet


class Datasheet:
    def __init__(
        self, workbook: xlsxwriter.Workbook, worksheet: xlsxwriter.worksheet.Worksheet
    ):
        self._args = {
            "column_width": 64,
            "header_height": 90,
            "supheader_height": 30,
            "border_weight": 2,
            "log2_tag": "[log2]",
            "nan_symbol": "n.a.",
            "sample_extraction_tag": "Intensity",
            "append_remaining_columns": False,
        }

        self.workbook = workbook
        self.worksheet = worksheet
        self.template_sections = None
        self._table = None
        self._sample_groups = None
        self._samples = []
        self._format_templates = {"default": {"align": "left", "num_format": "0"}}
        self._workbook_formats = {}
        self._conditional_formats = {}

    def apply_configuration(self, config_file: str) -> None:
        """Reads a config file and prepares workbook formats."""
        template = ReportTemplate.load(config_file)
        self.template_sections = template.sections
        self._args.update(template.settings)

        self._add_formats(template.formats)
        self._extend_header_format(template.sections)
        self._extend_supheader_format(template.sections)
        self._extend_border_formats()
        self._add_format_templates_to_workbook()

        self._add_conditional_formats(template.conditional_formats)

    def add_data(self, table: pd.DataFrame) -> None:
        """Adds table that will be used for filing the worksheet with data.

        Also extracts sample names from table by using the
        'sample_extraction_tag' from the config.
        """
        self._table = table.copy()
        # Replaces NaN in string columns with an empty string
        str_cols = self._table.select_dtypes(include=["object"]).columns
        self._table.loc[:, str_cols] = self._table.loc[:, str_cols].fillna("")

        extraction_tag = self._args["sample_extraction_tag"]
        if not self._samples and extraction_tag is not None:
            self._samples = _extract_samples_with_column_tag(
                self._table, extraction_tag
            )

    def write_data(self) -> None:
        """Writes data to the excel sheet and applies formatting."""
        if self.template_sections is None:
            raise ValueError(
                "Configuration has not applied. Call "
                '"ReportSheet.apply_configuration()" to do so.'
            )
        if self._table is None:
            raise ValueError(
                "No data for writing has been added. "
                'Call "ReportSheet.add_data()" to add data.'
            )

        # initialize data group writing
        data_groups = self._prepare_data_groups()
        hide_remaining_columns_group = False
        if self._args["append_remaining_columns"]:
            remaining_columns_data_group = self._prepare_remaining_columns_group()
            if remaining_columns_data_group is not None:
                hide_remaining_columns_group = True
                data_groups.append(remaining_columns_data_group)

        coordinates = {
            "supheader_row": 0,
            "header_row": 1,
            "data_row_start": 2,
            "data_row_end": 2,
            "first_column": 0,
            "last_column": None,
            "start_column": None,
        }
        coordinates["data_row_end"] += self._table.shape[0] - 1
        coordinates["last_column"] = coordinates["first_column"] - 1

        # write data groups
        for data_group in data_groups:
            coordinates["start_column"] = coordinates["last_column"] + 1
            coordinates["last_column"] = self._write_data_group(data_group, coordinates)

        # Hide the additionally added remaining columns
        if hide_remaining_columns_group:
            self.worksheet.set_column(
                coordinates["start_column"],
                coordinates["last_column"],
                options={"level": 1, "collapsed": True, "hidden": True},
            )

        # Set header height, add autofilter, freeze panes
        self.worksheet.set_row_pixels(
            coordinates["supheader_row"], self._args["supheader_height"]
        )
        self.worksheet.set_row_pixels(
            coordinates["header_row"], self._args["header_height"]
        )
        self.worksheet.autofilter(
            coordinates["data_row_start"] - 1,
            coordinates["first_column"],
            coordinates["data_row_end"],
            coordinates["last_column"],
        )
        self.worksheet.freeze_panes(coordinates["data_row_start"], 1)

    def _write_data_group(
        self, data_group: DataGroup, coordinates: dict[str, int]
    ) -> int:
        """Writes a data group to the excel sheet and applies formatting.

        Args:
            data_group:
            coordinates: Dicionary containing information about row and
                column positions. Keys are 'supheader_row', 'header_row',
                'data_row_start', 'data_row_end', 'start_column'.

        Returns:
            Position of last column that was added to the excel sheet
        """
        group_length = len(data_group.column_data)
        supheader_row = coordinates["supheader_row"]
        header_row = coordinates["header_row"]
        data_row_start = coordinates["data_row_start"]
        data_row_end = coordinates["data_row_end"]
        start_column = coordinates["start_column"]
        end_column = start_column + group_length - 1

        # Write supheader
        if data_group.supheader_text:
            excel_format = self.get_format(data_group.supheader_format)
            supheader_text = data_group.supheader_text
            if start_column != end_column:
                self.worksheet.merge_range(
                    supheader_row,
                    start_column,
                    supheader_row,
                    end_column,
                    supheader_text,
                    excel_format,
                )
            else:
                self.worksheet.write(
                    supheader_row, start_column, supheader_text, excel_format
                )

        # Write header data
        curr_column = start_column
        for text, format_name in zip(data_group.header_data, data_group.header_formats):
            excel_format = self.get_format(format_name)
            self.worksheet.write(header_row, curr_column, text, excel_format)
            curr_column += 1

        # Write column data
        curr_column = start_column
        for column_values, format_name in zip(
            data_group.column_data, data_group.column_formats
        ):
            excel_format = self.get_format(format_name)
            self.worksheet.write_column(
                data_row_start, curr_column, column_values, excel_format
            )
            curr_column += 1

        # Set column width
        self.worksheet.set_column_pixels(
            start_column, end_column, data_group.column_width
        )

        # Apply conditional formats
        for conditional_info in data_group.conditional_formats:
            format_name = conditional_info.name
            conditional_format = self.get_conditional(format_name)
            conditional_start = start_column + conditional_info.start
            conditional_end = start_column + conditional_info.end
            self.worksheet.conditional_format(
                data_row_start,
                conditional_start,
                data_row_end,
                conditional_end,
                conditional_format,
            )

        # Return last column position
        return end_column

    def _prepare_data_groups(self) -> list[DataGroup]:
        data_groups = []
        for group_name, group_config in self.template_sections.items():
            if _eval_arg("comparison_group", group_config):
                for group_data in self._prepare_comparison_group(
                    group_name, group_config
                ):
                    if group_data:
                        data_groups.append(group_data)
            elif _eval_arg("tag", group_config):
                group_data = self._prepare_sample_group(group_name, group_config)
                if group_data:
                    data_groups.append(group_data)
            else:
                group_data = self._prepare_data_group(group_name, group_config)
                if group_data:
                    data_groups.append(group_data)
        return data_groups

    def _prepare_data_group(self, group_name: str, config: dict) -> Optional[DataGroup]:
        """Prepare data required to write a feature group."""
        columns = [col for col in config["columns"] if col in self._table]
        if columns:
            conditional_formats = []
            if _eval_arg("column_conditional", config):
                for column_pos, column in enumerate(columns):
                    conditional = config["column_conditional"].get(column, None)
                    if conditional is not None:
                        conditional_formats.append(
                            ConditionalFormatGroupInfo(
                                conditional, column_pos, column_pos
                            )
                        )
            supheader_format_name = (
                f"supheader_{group_name}" if group_name else "supheader"
            )
            data_group = DataGroup(
                self._prepare_column_data(config, columns),
                self._prepare_column_formats(config, columns),
                self._prepare_header_data(config, columns),
                self._prepare_header_formats(config, columns, group_name),
                config.get("width", self._args["column_width"]),
                supheader_text=self._prepare_supheader_text(config),
                supheader_format=supheader_format_name,
                conditional_formats=conditional_formats,
            )
        else:
            data_group = None
        # Remove used columns from the table
        self._table = self._table.drop(columns=columns)
        return data_group

    def _prepare_sample_group(
        self, group_name: str, config: dict
    ) -> Optional[DataGroup]:
        """Prepare data required to write a sample group."""
        non_sample_columns, sample_columns = self._find_sample_group_columns(
            config["tag"]
        )
        config["columns"] = sample_columns
        data_group = self._prepare_data_group(group_name, config)

        if config["columns"] and _eval_arg("conditional", config):
            conditional_formats = []
            start = 0
            end = len(config["columns"]) - 1
            conditional_formats.append(
                ConditionalFormatGroupInfo(config["conditional"], start, end)
            )
            data_group.conditional_formats.extend(conditional_formats)
        return data_group

    def _prepare_comparison_group(self, group_name: str, config: dict):
        """Defines subgroups from a comparison group and prepares DataGroup instances.

        Each comparison of two samples generates one subgroup. To find all subgroups,
        each entry of "columns" from the comparison group is used as a tag to collect
        columns from self._table, then the search string is removed from each column
        and, if the remainder contains the "tag" specified by the comparison group, it
        is used as a subgroup.
        """
        comparison_groups = []
        for column_tag in config["columns"]:
            for column in _find_columns(self._table, column_tag):
                comparison_group = column.replace(column_tag, "").strip()
                if (
                    comparison_group not in comparison_groups
                    and config["tag"] in comparison_group
                ):
                    comparison_groups.append(comparison_group)

        data_groups = []
        for comparison_group in comparison_groups:
            matched = []
            for column in _find_columns(self._table, comparison_group):
                leftover = column.replace(comparison_group, "")
                for column_tag in config["columns"]:
                    leftover = leftover.replace(column_tag, "")
                if leftover.strip(WHITESPACE_SYMBOLS) == "":
                    matched.append(column)

            # Sort columns according to the order specified in the config file
            columns = []
            for column_tag in config["columns"]:
                columns.extend([col for col in matched if column_tag in col])

            # Define the sample comparison as supheader
            supheader = comparison_group
            if _eval_arg("replace_comparison_tag", config):
                supheader = supheader.replace(
                    config["tag"], config["replace_comparison_tag"]
                )

            # Prepare new config file for each comparison group
            sub_config = config.copy()
            sub_config["columns"] = columns
            sub_config["supheader"] = supheader
            sub_config["tag"] = comparison_group
            sub_config["remove_tag"] = True
            sub_config["column_conditional"] = {}
            for tag, conditional in config["column_conditional"].items():
                match = None
                for column in columns:
                    if column.find(tag) != -1:
                        match = column
                if match is not None:
                    sub_config["column_conditional"][match] = conditional
            data_group = self._prepare_data_group(group_name, sub_config)
            data_groups.append(data_group)
        return data_groups

    def _prepare_remaining_columns_group(self) -> Optional[DataGroup]:
        """Returns a remaining column data group or None, if no columns remain."""
        config = {"format": "default", "columns": self._table.columns.tolist()}
        data_group = self._prepare_data_group("", config)
        return data_group

    def _prepare_column_data(self, config: dict, columns: list[str]) -> dict:
        column_data = []
        for column in columns:
            values = self._table[column]
            if _eval_arg("log2", config):
                values = values.replace(0, np.nan)
                if not _intensities_in_logspace(values):
                    values = np.log2(values)
            values = values.replace(np.nan, self._args["nan_symbol"])
            column_data.append(values)
        return column_data

    def _prepare_column_formats(self, config: dict, columns: list[str]) -> dict:
        column_formats = []
        for column in columns:
            format_name = config["format"]
            if "column_format" in config and column in config["column_format"]:
                format_name = config["column_format"][column]
            column_formats.append(format_name)
        if _eval_arg("border", config):
            column_formats[0] = _rename_border_format(column_formats[0], left=True)
            column_formats[-1] = _rename_border_format(column_formats[-1], right=True)
        return column_formats

    def _prepare_header_data(self, config: dict, columns: list[str]) -> dict:
        header_data = []
        for text in columns:
            if _eval_arg("remove_tag", config):
                text = text.replace(config["tag"], "").strip()
            elif _eval_arg("log2", config):
                log2_tag = self._args["log2_tag"]
                text = f"{text} {log2_tag}".strip()
            header_data.append(text)
        return header_data

    def _prepare_header_formats(
        self, config: dict, columns: list[str], name: str
    ) -> dict:
        format_name = f"header_{name}" if name else "header"
        header_formats = [format_name for _ in columns]
        if _eval_arg("border", config):
            header_formats[0] = _rename_border_format(header_formats[0], left=True)
            header_formats[-1] = _rename_border_format(header_formats[-1], right=True)
        return header_formats

    def _prepare_supheader_text(self, config: dict) -> dict:
        text = config.get("supheader", None)
        if text and _eval_arg("log2", config):
            log2_tag = self._args["log2_tag"]
            text = f"{text} {log2_tag}".strip()
        return text

    def get_format(self, format_name: str) -> xlsxwriter.format.Format:
        """Returns an excel format."""
        return self._workbook_formats[format_name]

    def get_conditional(self, format_name: str) -> dict[str, object]:
        """Returns an excel conditional format."""
        return self._conditional_formats[format_name]

    def _add_args(self, args: dict[str, object]) -> None:
        """Add args from config file"""
        self._args.update(args)

    def _add_formats(self, formats: dict[str, dict[str, object]]) -> None:
        """Add formats."""
        for format_name in formats:
            format_properties = formats[format_name].copy()
            self._format_templates[format_name] = format_properties

    def _extend_header_format(self, groups: dict[str, object]) -> None:
        """Adds individual header formats per group.

        This allows to individualize header formats, such as defining a
        different background color. The default 'header' format is extended
        and modified by all entries from the groups 'header_format' entry.
        """
        self._extend_formats("header", groups)

    def _extend_supheader_format(self, groups: dict[str, object]) -> None:
        """Adds individual supheader formats per group.

        This allows to individualize supheader formats, such as defining a
        different background color. The default 'supheader' format is extended
        and modified by all entries from the groups 'supheader_format' entry.
        """
        self._extend_formats("supheader", groups)

    def _extend_formats(self, key: str, groups: dict[str, object]) -> None:
        """Adds individual format types per group.

        This allows to individualize header or supheader formats, such as
        defining a different background color or vertical rotation.
        The default format is extended and modified by all entries from the
        groups 'KEY_format' entry.
        """
        for group_name, group_info in groups.items():
            base_format = self._format_templates[key]
            group_format = base_format.copy()
            if f"{key}_format" in group_info:
                group_format.update(group_info[f"{key}_format"])
            group_format_name = f"{key}_{group_name}"
            self._format_templates[group_format_name] = group_format

    def _extend_border_formats(self) -> None:
        """Add format variants with borders to the format templates.

        For each format adds a variant with a left or a right border, the
        format name is extended by the _rename_border_format() function.
        """
        for name in list(self._format_templates):
            for left, right in [(True, False), (False, True), (True, True)]:
                format_name = _rename_border_format(name, left=left, right=right)
                format_properties = self._format_templates[name].copy()
                directions = [i for i, j in zip(["left", "right"], [left, right]) if j]
                for direction in directions:
                    format_properties[direction] = self._args["border_weight"]
                self._format_templates[format_name] = format_properties

    def _add_format_templates_to_workbook(self) -> None:
        """Add the template formats to the workbook."""
        for name, properties in self._format_templates.items():
            self._workbook_formats[name] = self.workbook.add_format(properties)

    def _add_conditional_formats(self, formats: dict[str, dict[str, object]]) -> None:
        """Add conditional formats to the conditional templates."""
        for format_name, format_properties in formats.items():
            self._conditional_formats[format_name] = format_properties

    def _find_sample_group_columns(self, tag: str) -> [list[str], list[str]]:
        """Find all columns belonging to a group, i.e. containing the 'tag'.

        Columns that contain the tag but not any of the extracted samples are
        added to the first list, columns that contain both the tag and a sample
        name are added to the second list.

        Returns:
            (list with non-sample columns, list with sample columns)
        """
        matched_columns = _find_columns(self._table, tag)
        non_sample_columns = []
        sample_columns = []
        for column in matched_columns:
            sample_query = column.replace(tag, "").strip()
            if sample_query in self._samples:
                sample_columns.append(column)
            else:
                non_sample_columns.append(column)
        return (non_sample_columns, sample_columns)


@dataclasses.dataclass
class DataGroup:
    """Class for storing data that will be written to the excel file.

    Attributes:
        column_data: An iterable of column values. Each column value entry must have
            the same length.
        column_formats: An iterable containing format names for each column.
        header_data: An iterable containing column titles.
        header_formats: An iterable containing format names for each header entry
        supheader_text: The column super header text.
        supheader_format: The format name of the column super header.
        width: Column width in pixel
        conditional_formats: An iterable containing

    The entry length of column_data, column_formats, header_data, header_formats has to
    be equal.
    """

    column_data: Iterable[Iterable[Union[str, int, float]]]
    column_formats: Iterable[str]
    header_data: Iterable[str]
    header_formats: Iterable[str]
    column_width: Union[int, float]
    supheader_text: str = ""
    supheader_format: str = ""
    conditional_formats: Iterable[ConditionalFormatGroupInfo] = dataclasses.field(
        default_factory=list
    )


@dataclasses.dataclass
class ConditionalFormatGroupInfo:
    """Class that stores a conditional format with start and end column positions."""

    name: str
    start: int
    end: int


def _create_empty_data_group() -> DataGroup:
    column_width = 45
    return DataGroup([[]], [], [], [], column_width)


def _extract_config_entry(config: dict[str, dict], name: str) -> dict[str, object]:
    return config.pop(name) if name in config else dict()


def _eval_arg(arg: str, args: dict) -> bool:
    """Evaluates wheter arg is present in args and is not False."""
    return arg in args and args[arg] is not False


def _extract_samples_with_column_tag(table: pd.DataFrame, tag: str) -> list[str]:
    """Extract sample names from columns containing the specified tag"""
    columns = _find_columns(table, tag, must_be_substring=True)
    samples = [c.replace(tag, "").strip() for c in columns]
    return samples


def _find_columns(
    df: pd.DataFrame, substring: str, must_be_substring: bool = False
) -> list[str]:
    """Returns a list of column names containing the substring.

    Args:
        df: Columns of this pandas.DataFrame are queried.
        substring: String that must be part of column names.
        must_be_substring: If true than column names are not reported if they
            are exactly equal to the substring.

    Returns:
        A list of column names
    """
    matches = [substring in col for col in df.columns]
    matched_columns = np.array(df.columns)[matches].tolist()
    if must_be_substring:
        matched_columns = [col for col in matched_columns if col != substring]
    return matched_columns


def _intensities_in_logspace(data: Union[pd.DataFrame, np.ndarray, Iterable]) -> bool:
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


def _rename_border_format(
    format_name: str, left: bool = False, right: bool = False
) -> str:
    """Adds a 'border tag' to the end of a format name.

    Args:
        left: If true, expands the format_name with the tag for the left border.
        right: If true, expands the format_name with the tag for the right border.

    Returns:
        The format_name with containing no, one or both border tags.
    """
    if left:
        format_name = f"{format_name}_left"
    if right:
        format_name = f"{format_name}_right"
    return format_name
