from __future__ import annotations
from dataclasses import dataclass
from typing import Optional, Protocol
import pandas as pd
import xlsxwriter  # type: ignore

from xlsxreport.compiler import prepare_compiled_sections
from xlsxreport.template import TableTemplate
from xlsxreport.writer import SectionWriter


class AbstractTabInfo(Protocol):
    """Abstract class for storing tab information."""

    tab_description: str
    tab_color: Optional[str]
    add_to_toc: bool


@dataclass
class TabInfo:
    """Class storing tab information."""

    tab_description: str
    tab_color: Optional[str]
    add_to_toc: bool


class AbstractTabWriter(Protocol):
    """Abstract class for writing content to a tab of an Excel file."""

    def write(self, workbook: xlsxwriter.Workbook, worksheet: xlsxwriter.Worksheet): ...


class AbstractTocWriter(Protocol):
    """Abstract class for writing a table of contents to a tab of an Excel file."""

    def set_tab_descriptions(
        self, tab_descriptions: dict[str, AbstractTabInfo]
    ) -> None: ...

    def write(self, workbook: xlsxwriter.Workbook, worksheet: xlsxwriter.Worksheet): ...


class ReportBuilder:
    """Class for building a multi-tab Excel report.

    To create a multi-tab Excel report, multiple tab writers can be added to the
    ReportBuilder successively. Each tab writer is responsible for writing the content
    of a tab to the Excel file. Once all tab writers have been added, calling the
    `build` method will create a workbook and add the report tabs. After building the
    report, the `close` method needs to be called to write the Excel file.
    Alternatively, the ReportBuilder can be used as a context manager, which will
    automatically build and close the report when the context is exited.
    """

    workbook: Optional[xlsxwriter.Workbook]
    filepath: str
    _tab_descriptions: dict[str, AbstractTabInfo]
    _tab_names: list[str]
    _tab_writers: dict[str, AbstractTabWriter]
    _toc_writers: dict[str, AbstractTocWriter]
    _built: bool

    def __init__(self, filepath: str):
        """Initialize the ReportBuilder instance.

        Args:
            filepath: The path of the Excel file where the report will be written.
        """
        self.workbook = None
        self.filepath = filepath
        self._tab_descriptions = {}
        self._tab_names = []
        self._tab_writers = {}
        self._toc_writers = {}
        self._report_built = False

    def __enter__(self) -> ReportBuilder:
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.build()
        self.close()

    def build(self) -> None:
        """Open a Workbook object and write the report tabs to the Workbook.

        This method needs to be called after all tab writers have been added to the
        ReportBuilder to open a Workbook object and add the report tabs. To write the
        Excel file and close the Workbook object, the `close` method needs to be called
        after building the report. If this method is called after the report has already
        been built, a ValueError is raised.
        """
        if self._report_built:
            raise ValueError("Report has already been built")

        self.workbook = xlsxwriter.Workbook(self.filepath)
        for tab_name in self._tab_names:
            worksheet = self.workbook.add_worksheet(tab_name)
            tab_info = self._tab_descriptions[tab_name]
            if tab_info.tab_color is not None:
                worksheet.set_tab_color(tab_info.tab_color)

        for tab_name, toc_writer in self._toc_writers.items():
            worksheet = self.workbook.get_worksheet_by_name(tab_name)
            toc_writer.set_tab_descriptions(self._tab_descriptions)
            toc_writer.write(self.workbook, worksheet)

        for tab_name, tab_writer in self._tab_writers.items():
            worksheet = self.workbook.get_worksheet_by_name(tab_name)
            tab_writer.write(self.workbook, worksheet)

        self._report_built = True

    def close(self) -> None:
        """Close the report and write the Excel file.

        This method needs to be called after the report has been built to write the
        Excel file. After calling this method, the Workbook is closed and cannot be
        accessed anymore. If this method is called before the report has been built or
        after the report has been closed, a ValueError is raised.
        """
        if not self._report_built or self.workbook is None:
            raise ValueError("Workbook has not been created, call build method first")
        self.workbook.close()

        self._report_built = False
        self._workbook = None

    def add_report_table(
        self,
        table: pd.DataFrame,
        table_template: TableTemplate,
        tab_name: str,
        tab_description: str = "",
        tab_color: Optional[str] = None,
        add_to_toc: bool = True,
    ) -> TableTemplate:
        """Add a tab to the Excel report containing a table formatted with xlsxreport.

        Args:
            table: The table to be added to the tab.
            table_template: The xlsxreport table template used for compiling the table
                sections that will be written to the tab.
            tab_name: The name of the Excel tab, must be unique and follow Excel tab
                naming rules.
            tab_description: The description of the tab that will be added to the table
                of contents, default "".
            tab_color: Optional, allows specifying a tab color for the Excel file. Must
                be a valid hex color code.
            add_to_toc: Whether to add the tab to a table of contents, default True.

        Returns:
            The table template used for compiling the table sections that will be
            written to the Excel file tab.
        """
        self.add_tab_writer(
            ReportTableWriter(table, table_template),
            tab_name,
            tab_description,
            tab_color,
            add_to_toc,
        )
        return table_template

    def add_table(
        self,
        table: pd.DataFrame,
        tab_name: str,
        tab_description: str = "",
        add_to_toc: bool = True,
        tab_color: Optional[str] = None,
    ) -> TableTemplate:
        """Add a tab to the Excel report containing a table with formatted headers.

        Args:
            table: The table to be added to the tab.
            tab_name: The name of the Excel tab, must be unique and follow Excel tab
                naming rules.
            tab_description: The description of the tab that will be added to the table
                of contents, default "".
            add_to_toc: Whether to add the tab to a table of contents, default True.
            tab_color: Optional, allows specifying a tab color for the Excel file. Must
                be a valid hex color code.

        Returns:
            The table template used for compiling the table sections that will be
            written to the Excel file tab.
        """
        formats = {
            "header": {
                "bold": True,
                "align": "left",
                "valign": "vcenter",
                "text_wrap": True,
                "bottom": 2,
                "top": 2,
            },
        }
        sections = {"all_columns": {"tag": "^."}}
        settings = {"freeze_cols": 0}
        table_template = TableTemplate.from_dict(
            {"formats": formats, "sections": sections, "settings": settings}
        )
        self.add_report_table(
            table, table_template, tab_name, tab_description, tab_color, add_to_toc
        )
        return table_template

    def add_toc(
        self, tab_name: str = "TOC", tab_description: str = "Table of content"
    ) -> None:
        """Add a table of contents (TOC) tab to the Excel report.

        args:
            tab_name: The name of the Excel tab, must be unique and follow Excel tab
                naming rules, default "TOC".
            tab_description: The description of the tab that will be added to the table
                of contents, default "Table of content".
        """
        self.add_toc_writer(TocWriter(), tab_name, tab_description=tab_description)

    def add_tab_writer(
        self,
        writer: AbstractTabWriter,
        tab_name: str,
        tab_description: str = "",
        tab_color: Optional[str] = None,
        add_to_toc: bool = True,
    ) -> None:
        """Add a tab writer instance to the ReportBuilder.

        This method adds a tab writer instance to the ReportBuilder, which is
        responsible for writing the content of a tab to the Excel file. The `write`
        method of the tab writer is called during the report building process to write
        the tab content to the specified tab of the Excel file.

        Args:
            writer: A tab writer instance that will be used to write the tab content.
            tab_name: The name of the Excel tab, must be unique and follow Excel tab
                naming rules.
            tab_description: The description of the tab that will be added to the table
                of contents, default "".
            tab_color: Optional, allows specifying a tab color for the Excel file. Must
                be a valid hex color code.
            add_to_toc: Whether to add the tab to a table of contents, default True.
        """
        self._add_tab_name(tab_name)
        self._tab_descriptions[tab_name] = TabInfo(
            tab_description, tab_color, add_to_toc
        )
        self._tab_writers[tab_name] = writer

    def add_toc_writer(
        self,
        writer: AbstractTocWriter,
        tab_name: str,
        tab_description: str = "",
        tab_color: Optional[str] = None,
        add_to_toc: bool = True,
    ) -> None:
        """Add a table of contents (TOC) writer instance to the ReportBuilder.

        This method adds a TOC writer instance to the ReportBuilder, which is
        responsible for writing a table of contents tab in the Excel report. The
        `set_tab_descriptions` method of the TOC writer is called first during the
        report building process to pass the descriptions of the report tabs.
        Subsequently, the `write` method of the TOC writer is called to write the table
        of contents to the specified tab of the Excel file.

        args:
            writer: The toc writer class that will be used to write the table of
                content.
            tab_name: The name of the Excel tab, must be unique and follow Excel tab
                naming rules.
            tab_description: The description of the tab that will be added to the table
                of contents, default "".
            tab_color: Optional, allows specifying a tab color for the Excel file. Must
                be a valid hex color code.
            add_to_toc: Whether to add the tab to a table of contents, default True.
        """
        self._add_tab_name(tab_name)
        self._tab_descriptions[tab_name] = TabInfo(
            tab_description, tab_color, add_to_toc
        )
        self._toc_writers[tab_name] = writer

    def _add_tab_name(self, tab_name: str) -> None:
        """Add a tab name to the ReportBuilder after checking if the tab name is valid.

        During the report build process, the added tab names are used to create the tabs
        in the Excel file.

        Args:
            tab_name: The name of the Excel tab, must be unique and follow Excel tab
                naming rules.
        """
        if tab_name in self._tab_names:
            raise ValueError(
                f"Tab name '{tab_name}' is already used, all tab names must be unique"
            )
        _validate_tab_name(tab_name)
        self._tab_names.append(tab_name)


class ReportTableWriter:
    """Class for writing a table formatted with xlsxreport to an Excel tab."""

    def __init__(
        self,
        table: pd.DataFrame,
        table_template: TableTemplate,
    ):
        self._table = table
        self._table_template = table_template

    def write(
        self,
        workbook: xlsxwriter.Workbook,
        worksheet: xlsxwriter.Worksheet,
    ) -> None:
        """Write the formatted table to the Excel file.

        Args:
            workbook: The xlsxwriter.Workbook instance that represents the Excel file.
            worksheet: The xlsxwriter.Worksheet instance that represents the tab where
                the table will be written.
        """
        compiled_sections = prepare_compiled_sections(self._table_template, self._table)
        SectionWriter(workbook).write_sections(
            worksheet, compiled_sections, self._table_template.settings
        )


class TocWriter:
    """Class for writing a Table of Contents to an Excel tab.

    Attributes:
        tab_descriptions: A dictionary with tab names as keys and `TabInfo`
            instances as values. The `TabInfo` instances contain information about
            the tab description, and whether the tab should be added to the TOC.
    """

    def __init__(self):
        self.tab_descriptions = {}

    def set_tab_descriptions(self, tab_descriptions: dict[str, AbstractTabInfo]):
        """Set the tab descriptions that are used for writing the table of contents."""
        self.tab_descriptions = tab_descriptions

    def write(
        self,
        workbook: xlsxwriter.Workbook,
        worksheet: xlsxwriter.Worksheet,
    ) -> None:
        """Write the table of contents (TOC) to the Excel file.

        Args:
            workbook: The xlsxwriter.Workbook instance that represents the Excel file.
            worksheet: The xlsxwriter.Worksheet instance that represents the tab where
                the table of content will be written.
        """
        name_column_width = 250
        description_column_width = 650

        worksheet.set_column_pixels(0, 0, name_column_width)
        worksheet.set_column_pixels(1, 1, description_column_width)
        _write_toc(workbook, worksheet, self.tab_descriptions)


def _write_toc(
    workbook: xlsxwriter.Workbook,
    worksheet: xlsxwriter.Worksheet,
    tab_descriptions: dict[str, AbstractTabInfo],
    first_row: int = 0,
) -> None:
    """Write a table of content to an Excel worksheet."""
    header_format = workbook.add_format(
        {
            "font_size": 14,
            "bold": True,
            "bg_color": "#d9d9d9",
            "top": 2,
            "bottom": 2,
            "left": 2,
            "right": 2,
        }
    )
    description_format = workbook.add_format({"text_wrap": True, "right": 2})
    name_format = workbook.add_format(
        {
            "text_wrap": True,
            "font_color": "#007F96",
            "underline": 1,
            "valign": "vcenter",
            "left": 2,
        }
    )
    bottom_format = workbook.add_format({"valign": "vcenter", "top": 2})

    worksheet.merge_range(first_row, 0, first_row, 1, "Table of content", header_format)
    for row, (tab_name, tab_info) in enumerate(tab_descriptions.items(), first_row + 1):
        if not tab_info.add_to_toc:
            continue
        worksheet.write_url(
            row, 0, f"internal:'{tab_name}'!A1", name_format, string=tab_name
        )
        worksheet.write(row, 1, tab_info.tab_description, description_format)
    worksheet.write_blank(row + 1, 0, None, bottom_format)
    worksheet.write_blank(row + 1, 1, None, bottom_format)


def _validate_tab_name(tab_name: str):
    """Validate the tab name according to Excel rules.

    Excel tab name rules:
    - The symbols [ ] : * ? / \ are not allowed in tab names
    - The tab name must be less than 32 characters.
    - The tab name cannot begin or end with an apostrophe.
    - Excel reserved tab name “History” is forbidden, also case insensitive variants
      such as “history” or “HISTORY”.
    """
    if len(tab_name) > 31:
        raise ValueError(
            f"Tab name '{tab_name}' is too long, must be less than 32 characters"
        )
    if any([i in tab_name for i in ["[", "]", ":", "*", "?", "/", "\\"]]):
        raise ValueError(
            f"Tab name '{tab_name}' contains invalid characters: [ ] : * ? / \\"
        )
    if tab_name.startswith("'") or tab_name.endswith("'"):
        raise ValueError(
            f"Tab name '{tab_name}' cannot begin or end with an apostrophe"
        )
    if tab_name.lower() == "history":
        raise ValueError(f"Tab name '{tab_name}' is a reserved name")
