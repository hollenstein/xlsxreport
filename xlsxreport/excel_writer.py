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
        start_column: int = 0,
        start_row: int = 0,
    ):
        """Write a list of sections to the workbook to create a table."""
        sections = sections if sections is not None else {}
        # column_height = settings["column_height"]  # -> NOVEL
        # header_height = settings["header_height"]
        # supheader_height = settings["supheader_height"]

        # 1) setup coordinates (take into account if a supheader should be written)
        # 2) write sections
        # 3) set supheader, header and row heights (needs to know column lengths)
        # 4) freeze panes
        # 5) add auto filter

    def _write_section(
        self,
        worksheet: xlsxwriter.worksheet.Worksheet,
        section: TableSection,
        start_row: int,
        start_column: int,
        write_supheader: bool,
    ):
        """Write a TableSection to the workbook."""
        header_row = start_row
        if write_supheader:
            header_row += 1
        for column_num, column in enumerate(section.data.columns):
            self._write_column(
                worksheet,
                row=header_row,
                column=start_column + column_num,
                header=section.headers[column],
                values=section.data[column],
                header_format=section.header_formats[column],
                values_format=section.column_formats[column],
                conditional_format=section.column_conditionals[column],
                column_width=section.column_widths[column],
            )
        # 1) if write_supheader: merge_range -> text and format; move header_row
        #   section.supheader
        #   section.supheader_format
        # 2) write section conditional format: conditional_format
        #   section.section_conditional

    def _write_column(
        self,
        worksheet: xlsxwriter.worksheet.Worksheet,
        row: int,
        column: int,
        header: str,
        values: Iterable,
        header_format: xlsxwriter.format.Format,
        values_format: xlsxwriter.format.Format,
        conditional_format: dict[str, float | str | bool],
        column_width: float,
    ):
        header_xlsx_format = self.get_xlsx_format(header_format)
        values_xlsx_format = self.get_xlsx_format(values_format)
        worksheet.write(row, column, header, header_xlsx_format)
        worksheet.write_column(row + 1, column, values, values_xlsx_format)
        worksheet.set_column_pixels(column, column, column_width)
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
