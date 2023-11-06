from __future__ import annotations
from typing import Iterable
import xlsxwriter

from xlsxreport.compiler import TableSection


class TableSectionWriter:
    def __init__(self, workbook: xlsxwriter.Workbook):
        self.workbook = workbook
        self._xlsxwriter_formats: dict = {}  # use dictionary hash as key

    def write_sections(
        self,
        worksheet: xlsxwriter.worksheet.Worksheet,
        sections: Iterable[TableSection],
        settings: dict | None = None,
        start_row: int = 0,
        start_column: int = 0,
    ) -> None:
        """Write a list of sections to the workbook to create a table."""
        settings = settings if settings is not None else {}
        write_supheader = settings.get("write_supheader", True)
        supheader_height = settings.get("supheader_height", 20)
        header_height = settings.get("header_height", 20)

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

        # TODO - not tested from here on (including last value row)
        if write_supheader:
            worksheet.set_row_pixels(start_row, supheader_height)
        worksheet.set_row_pixels(header_row, header_height)
        worksheet.freeze_panes(values_row, start_column + 1)
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
            num_values, num_rows = section.data.shape
            worksheet.conditional_format(
                values_row,
                start_column,
                values_row + num_values - 1,
                start_column + num_rows - 1,
                section.section_conditional,
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
        if not supheader:
            return
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
        values: Iterable,
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
        """Converts a format description to an xlsxwriter format.

        Args:
            format_description: A dictionary describing the format.
        """
        _hash = _hashable_from_dict(format_description)
        if _hash not in self._xlsxwriter_formats:
            self._xlsxwriter_formats[_hash] = self.workbook.add_format(
                format_description
            )
        return self._xlsxwriter_formats[_hash]


def _hashable_from_dict(format_description: dict[str, float | str | bool]):
    return tuple((k, format_description[k]) for k in sorted(format_description))
