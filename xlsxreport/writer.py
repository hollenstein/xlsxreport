"""This module provides a class for writing compiled TableSections to an Excel file."""
from __future__ import annotations
from typing import Collection, Iterable, Optional, Protocol
import warnings

import pandas as pd
import xlsxwriter.format
import xlsxwriter.worksheet


class TableSection(Protocol):
    """Contains information for writing and formatting a section of a table."""

    data: pd.DataFrame
    column_formats: dict
    column_conditionals: dict
    column_widths: dict
    headers: dict
    header_formats: dict
    supheader: str
    supheader_format: dict
    section_conditional: dict
    hide_section: bool


class TableSectionWriter:
    """Class for writing compiled `TableSections` to an Excel file.

    The `TableSectionWriter` provides the method `write_sections` to write a list of
    compiled `TableSection`s to a worksheet in an Excel file.

    Attributes:
        workbook: The xlsxwriter.Workbook instance that represents the Excel file to
            which compiled `TableSection`s will be written.
    """

    def __init__(self, workbook: xlsxwriter.Workbook):
        """Initialize a TableSectionWriter.

        Args:
            workbook: The xlsxwriter.Workbook instance that represents the Excel file to
                which compiled `TableSection`s will be written.
        """
        self.workbook = workbook
        self._xlsxwriter_formats: dict = {}  # use dictionary hash as key

    def write_sections(
        self,
        worksheet: xlsxwriter.worksheet.Worksheet,
        sections: Iterable[TableSection],
        settings: Optional[dict] = None,
        start_row: int = 0,
        start_column: int = 0,
    ) -> None:
        """Write a list of compiled `TableSection`s to the `worksheet` in an Excel file.

        Args:
            worksheet: The Excel worksheet to write to.
            sections: A list of compiled `TableSection`s that will be written as a
                continuous table to the `worksheet`.
            settings: Optional, specify general settings for the table. The following
                settings are available: `write_supheader`, `supheader_height`, and
                `header_height`.
            start_row: The row in the Excel worksheet to start writing the sections at.
                The first row as seen in Excel starts at 0. The default is 0.
            start_column: The column in the Excel worksheet to start writing the
                sections at. The first column as seen in Excel starts at 0. The default
                is 0.
        """
        settings: dict = settings if settings is not None else {}
        write_supheader: bool = settings.get("write_supheader", False)
        supheader_height: float = settings.get("supheader_height", 20)
        header_height: float = settings.get("header_height", 20)
        add_autofiler: bool = settings.get("add_autofilter", True)
        freeze_cols: int = settings.get("freeze_cols", 1)

        header_row = start_row
        values_row = start_row + 1
        if write_supheader:
            header_row += 1
            values_row += 1
        next_column = start_column
        last_value_row = start_row

        for section in sections:
            self._write_section(
                worksheet=worksheet,
                section=section,
                start_row=start_row,
                start_column=next_column,
                write_supheader=write_supheader,
            )
            last_value_row = max(last_value_row, section.data.shape[0] + header_row)
            next_column += section.data.shape[1]

        # TODO - not tested from here on (including calculation of last_value_row)
        if write_supheader:
            worksheet.set_row_pixels(start_row, supheader_height)
        worksheet.set_row_pixels(header_row, header_height)
        if freeze_cols > 0:
            worksheet.freeze_panes(values_row, start_column + freeze_cols)
        if add_autofiler:
            worksheet.autofilter(
                header_row,
                start_column,
                last_value_row,
                last_col=next_column - 1,
            )

    def _write_section(
        self,
        worksheet: xlsxwriter.worksheet.Worksheet,
        section: TableSection,
        start_row: int,
        start_column: int,
        write_supheader: bool,
    ) -> None:
        """Write a TableSection to the workbook."""
        num_values, num_rows = section.data.shape
        header_row = start_row
        values_row = start_row + 1

        if write_supheader:
            header_row += 1
            values_row += 1
            self._write_supheader(
                worksheet=worksheet,
                row=start_row,
                column=start_column,
                num_columns=section.data.shape[1],
                supheader=section.supheader,
                supheader_format=section.supheader_format,
            )
        for column_position, column in enumerate(section.data.columns):
            self._write_column(
                worksheet=worksheet,
                row=header_row,
                column=start_column + column_position,
                header=section.headers[column],
                values=section.data[column],
                header_format=section.header_formats[column],
                values_format=section.column_formats[column],
                conditional_format=section.column_conditionals[column],
                column_width=section.column_widths[column],
            )
        if section.section_conditional:
            worksheet.conditional_format(
                values_row,
                start_column,
                values_row + num_values - 1,
                start_column + num_rows - 1,
                section.section_conditional,
            )
        if section.hide_section:
            worksheet.set_column(
                start_column,
                start_column + num_rows - 1,
                options={"level": 1, "collapsed": True, "hidden": True},
            )

    def _write_supheader(
        self,
        worksheet: xlsxwriter.worksheet.Worksheet,
        row: int,
        column: int,
        num_columns: int,
        supheader: str,
        supheader_format: dict[str, float | str | bool],
    ) -> None:
        """Write a supheader to the workbook by merging a range of cells."""
        supheader_xlsx_format = self.get_xlsx_format(supheader_format)
        # if not supheader:
        #    return
        if num_columns > 1:
            last_column = column + num_columns - 1
            worksheet.merge_range(
                row, column, row, last_column, supheader, supheader_xlsx_format
            )
        else:
            worksheet.write(row, column, supheader, supheader_xlsx_format)

    def _write_column(
        self,
        worksheet: xlsxwriter.worksheet.Worksheet,
        row: int,
        column: int,
        header: str,
        values: Collection,
        header_format: dict[str, float | str | bool],
        values_format: dict[str, float | str | bool],
        conditional_format: dict[str, float | str | bool],
        column_width: float,
    ) -> None:
        """Write a column to the workbook."""
        header_xlsx_format = self.get_xlsx_format(header_format)
        values_xlsx_format = self.get_xlsx_format(values_format)
        worksheet.write(row, column, header, header_xlsx_format)
        worksheet.write_column(row + 1, column, values, values_xlsx_format)
        worksheet.set_column_pixels(column, column, column_width)
        if conditional_format:
            worksheet.conditional_format(
                row + 1, column, row + len(values), column, conditional_format
            )

    def get_xlsx_format(
        self, format_description: dict[str, float | str | bool]
    ) -> xlsxwriter.format.Format:
        """Converts a format description to an `xlsxwriter` format.

        Args:
            format_description: A dictionary describing the format.
        """
        _hash = _hashable_from_dict(format_description)
        if _hash not in self._xlsxwriter_formats:
            try:
                xlsx_format = self.workbook.add_format(format_description)
            except AttributeError:
                warnings.warn(
                    f"Invalid xlsxwriter format description: {format_description}",
                    UserWarning,
                )
                xlsx_format = self.workbook.add_format({})
            self._xlsxwriter_formats[_hash] = xlsx_format
        return self._xlsxwriter_formats[_hash]


def _hashable_from_dict(format_description: dict[str, float | str | bool]):
    return tuple((k, format_description[k]) for k in sorted(format_description))
