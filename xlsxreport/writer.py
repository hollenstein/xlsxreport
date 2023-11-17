from __future__ import annotations
from typing import Optional

import pandas as pd
import xlsxwriter

from xlsxreport.excel_writer import TableSectionWriter
from xlsxreport.template import ReportTemplate
from xlsxreport.compiler import prepare_table_sections


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
        self.workbook = workbook
        self.worksheet = worksheet
        self.section_writer = TableSectionWriter(self.workbook)
        self.table = None
        self.report_template = None

    def apply_configuration(self, config_file: str) -> None:
        """Reads a config file and prepares workbook formats."""
        self.report_template = ReportTemplate.load(config_file)
        self.report_template.settings["evaluate_log2_transformation"] = True

    def add_data(self, table: pd.DataFrame) -> None:
        self.table = table

    def write_data(self) -> None:
        table_sections = prepare_table_sections(self.report_template, self.table)
        self.section_writer.write_sections(
            self.worksheet, table_sections, settings=self.report_template.settings
        )
