from typing import Optional
import xlsxwriter

from xlsxreport.compiler import TableSection


class TableSectionWriter:
    def __init__(self, workbook: xlsxwriter.Workbook):
        self.workbook = workbook
        self._xlsxwriter_formats: dict = {}  # use dictionary hash as key

    def write_sections(
        self,
        sections: list[TableSection],
        settings: Optional[dict] = None,
        start_column: int = 0,
        start_row: int = 0,
    ):
        sections = sections if sections is not None else {}
        # column_height = settings["column_height"]  # -> NOVEL
        # header_height = settings["header_height"]
        # supheader_height = settings["supheader_height"]

    def _write_section(
        section: TableSection, start_column: int, start_row: int, write_supheader: bool
    ):
        ...

    def get_format(self, format_description: dict) -> xlsxwriter.format.Format:
        ...
