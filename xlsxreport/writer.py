"""
Description of YAML config file:
- One or multiple columns are always defined as a group
- The order of groups if specified by the order in the config file
- Each group has a name
- There are three types of groups: feature, tag, comparison
    - feature: directly specify a sequence of columns
    - tag: quantitative columns that all contain a specific tag, e.g. "Intensity"
    - comparison: Columns that contain a quantitative comparison between
        multiple experiments.

TODOs
- Add option to append all remaining columns (and hide them)
- Add option to specify sample order
    - Requires that samples are specified by user
    - Adapt _find_sample_group_columns() to sort columns
- Add column comments
- Add option to sort the table before writing data
"""
from typing import Iterable, Optional, Union

import numpy as np
import pandas as pd
import xlsxwriter
import yaml


class Reportbook(xlsxwriter.Workbook):
    """Subclass of the XlsxWriter Workbook class."""

    def add_infosheet(self):
        worksheet = self.add_worksheet("info")
        return worksheet

    def add_datasheet(self, name: Optional[str] = None):
        worksheet = self.add_worksheet(name)
        data_sheet = Datasheet(self, worksheet)
        return data_sheet


class Datasheet:
    def __init__(
        self, workbook: xlsxwriter.Workbook, worksheet: xlsxwriter.worksheet.Worksheet
    ):
        self._args = {
            "border_weight": 2,
            "log2_tag": "[log2]",
            "nan_symbol": "n.a.",
            "supheader_height": 30,
            "header_height": 90,
            "column_width": 64,
            "sample_extraction_tag": "Intensity",
        }

        self.workbook = workbook
        self.worksheet = worksheet
        self._config = None
        self._table = None
        self._sample_groups = None
        self._samples = []
        self._format_templates = {}
        self._workbook_formats = {}
        self._conditional_formats = {}

    def apply_configuration(self, config_file: str) -> None:
        """Reads a config file and prepares workbook formats."""
        self._config = parse_config_file(config_file)
        self._add_args(self._config["args"])

        self._add_formats(self._config["formats"])
        self._extend_header_format(self._config["groups"])
        self._extend_supheader_format(self._config["groups"])
        self._extend_border_formats()
        self._add_format_templates_to_workbook()

        self._add_conditional_formats(self._config["conditional_formats"])

    def add_data(self, table: pd.DataFrame) -> None:
        """Adds table that will be used for filing the worksheet with data.

        Also extracts sample names from table by using the
        'sample_extraction_tag' from the config.
        """
        self._table = table.copy()
        # Replace NaN in string columns with an empty string
        str_cols = self._table.select_dtypes(include=["object"]).columns
        self._table.loc[:, str_cols] = self._table.loc[:, str_cols].fillna("")

        extraction_tag = self._args["sample_extraction_tag"]
        if not self._samples and extraction_tag is not None:
            self._samples = _extract_samples_with_column_tag(
                self._table, extraction_tag
            )

    def write_data(self) -> None:
        """Writes data to the excel sheet and applies formatting."""
        if self._config is None:
            raise Exception(
                "Configuration has not applied. Call "
                '"ReportSheet.apply_configuration()" to do so.'
            )
        if self._table is None:
            raise Exception(
                "No data for writing has been added. "
                'Call "ReportSheet.add_data()" to add data.'
            )

        # initialize data group writing
        data_groups = self._prepare_data_groups()
        coordinates = {
            "supheader_row": 0,
            "header_row": 1,
            "data_row_start": 2,
            "data_row_end": 2,
            "first_column": 0,
        }
        coordinates["data_row_end"] += self._table.shape[0] - 1
        coordinates["start_column"] = coordinates["first_column"]

        # write data groups
        for data_group in data_groups:
            coordinates["last_column"] = self._write_data_group(data_group, coordinates)
            coordinates["start_column"] = coordinates["last_column"] + 1

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

    def _write_data_group(self, data_group: dict, coordinates: dict[str, int]) -> int:
        """Writes a data group to the excel sheet and applies formatting.

        Args:
            data_group:
            coordinates: Dicionary containing information about row and
                column positions. Keys are 'supheader_row', 'header_row',
                'data_row_start', 'data_row_end', 'start_column'.

        Returns:
            Last column position that was filled with data
        """
        group_length = len(data_group["data"])
        supheader_row = coordinates["supheader_row"]
        header_row = coordinates["header_row"]
        data_row_start = coordinates["data_row_start"]
        data_row_end = coordinates["data_row_end"]
        start_column = coordinates["start_column"]
        end_column = start_column + group_length - 1

        # Write column data
        curr_column = start_column
        for values, format_name, conditional_name in data_group["data"]:
            excel_format = self.get_format(format_name)
            self.worksheet.write_column(
                data_row_start, curr_column, values, excel_format
            )
            if conditional_name:
                excel_conditional = self.get_conditional(conditional_name)
                self.worksheet.conditional_format(
                    data_row_start,
                    curr_column,
                    data_row_end,
                    curr_column,
                    excel_conditional,
                )
            curr_column += 1

        # Write header data
        curr_column = start_column
        for text, format_name in data_group["header"]:
            excel_format = self.get_format(format_name)
            self.worksheet.write(header_row, curr_column, text, excel_format)
            curr_column += 1

        # Write supheader
        supheader_text, format_name = data_group["supheader"]
        if supheader_text:
            supheader_format = self.get_format(format_name)
            self.worksheet.merge_range(
                supheader_row,
                start_column,
                supheader_row,
                end_column,
                supheader_text,
                supheader_format,
            )

        # Set column width
        self.worksheet.set_column_pixels(
            start_column, end_column, data_group["col_width"]
        )

        # Apply conditional formats to the group
        for conditional_info in data_group["conditional_formats"]:
            format_name = conditional_info["name"]
            conditional_format = self.get_conditional(format_name)
            conditional_start = start_column + conditional_info["start"]
            conditional_end = start_column + conditional_info["end"]
            self.worksheet.conditional_format(
                data_row_start,
                conditional_start,
                data_row_end,
                conditional_end,
                conditional_format,
            )

        # Return last column
        return end_column

    def _prepare_data_groups(self):
        data_groups = []
        for name, config in self._config["groups"].items():
            if _eval_arg("comparison_group", config):
                for group_data in self._prepare_comparison_group(name, config):
                    if group_data:
                        data_groups.append(group_data)
            elif _eval_arg("tag", config):
                group_data = self._prepare_sample_group(name, config)
                if group_data:
                    data_groups.append(group_data)
            else:
                group_data = self._prepare_feature_group(name, config)
                if group_data:
                    data_groups.append(group_data)
        return data_groups

    def _prepare_feature_group(self, name, config) -> dict():
        """Prepare data required to write a feature group."""
        columns = [col for col in config["columns"] if col in self._table]
        if columns:
            group_data = {
                "data": self._prepare_column_data(config, columns),
                "header": self._prepare_column_headers(config, columns, name),
                "supheader": self._prepare_supheader(config, name),
                "col_width": config.get("width", self._args["column_width"]),
                "conditional_formats": [],
            }
            # Remove already used columns from the table
            self._table = self._table.drop(columns=columns)
        else:
            group_data = {}

        return group_data

    def _prepare_sample_group(self, name, config):
        """Prepare data required to write a sample group."""
        non_sample_columns, sample_columns = self._find_sample_group_columns(
            config["tag"]
        )
        conditional_formats = []
        end = -1
        for columns in [non_sample_columns, sample_columns]:
            if columns:
                start = end + 1
                end = start + len(columns) - 1
                conditional_formats.append(
                    {"name": config["conditional"], "start": start, "end": end}
                )

        columns = [*non_sample_columns, *sample_columns]
        if columns:
            group_data = {
                "data": self._prepare_column_data(config, columns),
                "header": self._prepare_column_headers(config, columns, name),
                "supheader": self._prepare_supheader(config, name),
                "col_width": config.get("width", self._args["column_width"]),
                "conditional_formats": conditional_formats,
            }
            # Remove already used columns from the table
            self._table = self._table.drop(columns=columns)
        else:
            group_data = {}

        return group_data

    def _prepare_comparison_group(self, name, config):
        # Find all comparison groups
        comparison_groups = []
        for column_tag in config["columns"]:
            for column in _find_columns(self._table, column_tag):
                comparison_group = column.replace(column_tag, "").strip()
                if comparison_group not in comparison_groups:
                    comparison_groups.append(comparison_group)

        comparison_group_data = []
        for comparison_group in comparison_groups:
            # Sort columns according to the order specified in the config file
            matched = _find_columns(self._table, comparison_group)
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
            sub_config["supheader"] = comparison_group
            sub_config["tag"] = comparison_group
            sub_config["column_conditional"] = {}
            for tag, conditional in config["column_conditional"].items():
                match = None
                for column in columns:
                    if column.find(tag) != -1:
                        match = column
                if match is not None:
                    sub_config["column_conditional"][match] = conditional
            group_data = self._prepare_feature_group(name, sub_config)
            comparison_group_data.append(group_data)
        return comparison_group_data

    def _prepare_column_data(self, config: dict, columns: list[str]) -> dict:
        data_info = []
        for column in columns:
            data = self._table[column]
            if _eval_arg("log2", config):
                data = data.replace(0, np.nan)
                if not _intensities_in_logspace(data):
                    data = np.log2(data)
            data = data.replace(np.nan, self._args["nan_symbol"])

            format_name = config["format"]
            conditional = None
            if "column_format" in config and column in config["column_format"]:
                format_name = config["column_format"][column]
            if "column_conditional" in config:
                conditional = config["column_conditional"].get(column, None)
            data_info.append([data, format_name, conditional])
        if _eval_arg("border", config):
            data_info[0][1] = f"{data_info[0][1]}_left"
            data_info[-1][1] = f"{data_info[-1][1]}_right"
        return data_info

    def _prepare_column_headers(
        self, config: dict, columns: list[str], name: str
    ) -> dict:
        header_info = []
        for text in columns:
            if _eval_arg("remove_tag", config):
                text = text.replace(config["tag"], "").strip()
            elif _eval_arg("log2", config):
                log2_tag = self._args["log2_tag"]
                text = f"{text} {log2_tag}".strip()
            format_name = f"header_{name}"
            header_info.append([text, format_name])
        if _eval_arg("border", config):
            header_info[0][1] = f"{header_info[0][1]}_left"
            header_info[-1][1] = f"{header_info[-1][1]}_right"
        return header_info

    def _prepare_supheader(self, config: dict, name: str) -> dict:
        text = config.get("supheader", None)
        if text and _eval_arg("log2", config):
            log2_tag = self._args["log2_tag"]
            text = f"{text} {log2_tag}".strip()
        format_name = f"supheader_{name}"
        supheader_info = [text, format_name]
        return supheader_info

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

    def _extend_border_formats(self) -> None:
        """Add format variants with borders to the format templates.

        For each format adds a variant with a left or a right border, the
        format name is extended by 'format_left' or 'format_right'.
        """
        for name in list(self._format_templates):
            for border in ["left", "right", "left_right"]:
                format_name = f"{name}_{border}"
                format_properties = self._format_templates[name].copy()
                for direction in border.split("_"):
                    format_properties[direction] = self._args["border_weight"]
                self._format_templates[format_name] = format_properties

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

    def _add_format_templates_to_workbook(self):
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


def parse_config_file(file: str) -> dict[str, dict]:
    """Parses excel report config file and returns entries as dictionaries.

    Returns:
        Dictionary containing the keys 'formats', 'conditional_formats',
            'groups', 'args', each pointing to another dictionary.
    """
    with open(file) as open_file:
        yaml_file = yaml.safe_load(open_file)
    config = {
        "args": _extract_config_entry(yaml_file, "args"),
        "groups": _extract_config_entry(yaml_file, "groups"),
        "formats": _extract_config_entry(yaml_file, "formats"),
        "conditional_formats": _extract_config_entry(yaml_file, "conditional_formats"),
    }
    return config


def _extract_config_entry(config: dict[str, dict], name: str) -> dict[str, object]:
    return config.pop(name) if name in config else dict()


def _eval_arg(arg: str, args: dict) -> bool:
    """Evaluates wheter arg is present in args and is not False."""
    return arg in args and args[arg] is not False


def _extract_samples_with_column_tag(table: pd.DataFrame, tag: str):
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
