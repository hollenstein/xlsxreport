from io import BytesIO
from unittest.mock import MagicMock, call

import openpyxl
import pandas as pd
import pytest
from xlsxwriter import Workbook as Workbook

import xlsxreport.compiler as compiler
import xlsxreport.excel_writer as writer


class ExcelWriteReadManager:
    """Helper class for writing an excel file with `xlsxwriter` to buffer and reading it
    from the buffer with `openpyxl`.

    Note that for writing rows and columns xlsxwriter is zero indexed whereas pyopenxl
    is one indexed.
    """

    def __init__(self):
        self.buffer = None
        self.loaded_workbook = None
        self.loaded_worksheet = None

    def __enter__(self):
        self.buffer = BytesIO()
        self.workbook = Workbook(self.buffer)
        self.worksheet_name = "Worksheet"
        self.worksheet = self.workbook.add_worksheet(self.worksheet_name)
        self.section_writer = writer.TableSectionWriter(self.workbook)
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.workbook.close()

    def load_excel_from_buffer(self):
        if self.buffer is None:
            raise Exception("Buffer is not initialized")
        self.loaded_workbook = openpyxl.load_workbook(self.buffer)
        self.loaded_worksheet = self.loaded_workbook[self.worksheet_name]

    def load_worksheet(self) -> openpyxl.worksheet.worksheet.Worksheet:
        self.load_excel_from_buffer()
        return self.loaded_worksheet


@pytest.fixture()
def table_section() -> compiler.TableSection:
    section = compiler.TableSection(
        data=pd.DataFrame({"Column 1": [1, 2, 3], "Column 2": ["A", "B", "C"]}),
        column_formats={"Column 1": {}, "Column 2": {}},
        column_conditionals={"Column 1": {}, "Column 2": {}},
        column_widths={"Column 1": 10, "Column 2": 10},
        headers={"Column 1": "Column 1", "Column 2": "Column 2"},
        header_formats={"Column 1": {}, "Column 2": {}},
        supheader="Supheader",
        supheader_format={},
        section_conditional={},
    )
    return section


class TestIntegrationTableSectionWriteColumn:
    @pytest.fixture(autouse=True)
    def _init_section_writer_with_write_column(self):
        # Note that xlsxwriter is zero indexed whereas pyopenxl is one indexed
        row, column = 2, 2
        with ExcelWriteReadManager() as excel_manager:
            excel_manager.section_writer._write_column(
                excel_manager.worksheet,
                row=row - 1,
                column=column - 1,
                header="Header",
                values=[1, 2, "3"],
                header_format={"bold": True, "bottom": 2},
                values_format={"align": "center", "num_format": "0.00"},
                conditional_format={"type": "2_color_scale"},
                column_width=10000,
            )
        self.worksheet = excel_manager.load_worksheet()
        self.written_column = list(self.worksheet.columns)[1]

    def test_column_is_written_to_correct_position(self):
        columns = list(self.worksheet.columns)
        empty_col_position = 0
        written_col_position = 1
        assert all([cell.value is None for cell in columns[empty_col_position]])
        assert any([cell.value is None for cell in columns[written_col_position]])

    def test_correct_column_values_written(self):
        column_values = [cell.value for cell in self.written_column]
        assert column_values == [None, "Header", 1, 2, "3"]

    def test_header_format_is_applied(self):
        header_cell = self.worksheet.cell(row=2, column=2)
        assert header_cell.font.bold == True
        assert header_cell.border.bottom.style == "medium"

    def test_column_format_is_applied(self):
        column_cells = self.written_column[2:]  # without empty and header cell
        assert all([cell.alignment.horizontal == "center" for cell in column_cells])
        assert all([cell.number_format == "0.00" for cell in column_cells])

    def test_set_correct_column_width(self):
        # xlsxwriter sets column width in pixel units, whereas openpyxl sets width in
        # some kind of "unit" format. So we set the width to a high number and check if
        # the column width in "units" is smaller or bigger than 100.
        empty_col_width = self.worksheet.column_dimensions["A"].width
        written_col_width = self.worksheet.column_dimensions["B"].width
        assert empty_col_width < 100
        assert written_col_width > 100

    def test_conditional_format_applied_to_correct_area(self):
        conditional_format = list(self.worksheet.conditional_formatting)[0]
        conditional_format.cells.ranges[0].top == [(3, 2)]
        conditional_format.cells.ranges[0].bottom == [(5, 2)]


class TestTableSectionWriteColumn:
    @pytest.fixture(autouse=True)
    def _init_section_writer(self):
        self.worksheet_mock = MagicMock(name="worksheet_mock")
        self.section_writer = writer.TableSectionWriter(Workbook())

    @pytest.fixture(autouse=True)
    def _init_arguments_for_write_column(self):
        self.args = {
            "row": 1,
            "column": 1,
            "header": "Header name",
            "values": [1, 2, 3],
            "header_format": {"bold": True},
            "values_format": {"bold": False},
            "conditional_format": {"type": "cell"},
            "column_width": 10,
        }
        self.header_row = self.args["row"]
        self.data_row_start = self.args["row"] + 1
        self.data_row_end = self.args["row"] + len(self.args["values"])

    def test_write_column_called_with_correct_arguments_to_write_values(self):
        self.section_writer._write_column(self.worksheet_mock, **self.args)
        assert self.worksheet_mock.write_column.call_args == call(
            self.data_row_start,
            self.args["column"],
            self.args["values"],
            self.section_writer.get_xlsx_format(self.args["values_format"]),
        )

    def test_write_called_with_correct_arguments_to_write_header(self):
        self.section_writer._write_column(self.worksheet_mock, **self.args)
        assert self.worksheet_mock.write.call_args == call(
            self.header_row,
            self.args["column"],
            self.args["header"],
            self.section_writer.get_xlsx_format(self.args["header_format"]),
        )

    def test_set_column_pixels_called_with_correct_arguments(self):
        self.section_writer._write_column(self.worksheet_mock, **self.args)
        assert self.worksheet_mock.set_column_pixels.call_args == call(
            self.args["column"], self.args["column"], self.args["column_width"]
        )

    def test_conditional_format_called_with_correct_arguments(self):
        self.section_writer._write_column(self.worksheet_mock, **self.args)

        assert self.worksheet_mock.conditional_format.call_args == call(
            self.data_row_start,
            self.args["column"],
            self.data_row_end,
            self.args["column"],
            self.args["conditional_format"],
        )

    def test_conditional_format_not_called_when_conditionaL_format_is_empty(self):
        self.args["conditional_format"] = {}
        self.section_writer._write_column(self.worksheet_mock, **self.args)
        self.worksheet_mock.conditional_format.assert_not_called()


class TestTableSectionWriteSupheader:
    def _create_worksheet_with_section_writer_and_write_supheader(
        self, row: int = 1, column: int = 1, num_columns: int = 1
    ):
        """Creates an xlsxwriter.workbook, an xlsxwriter.worksheet and a
        TableSectionWriter. Then writes the `table_section` to the worksheet by using
        the TableSectionWriter._write_section method. The worksheet is then safed to a
        buffer and loaded with openpyxl. The loaded worksheet is returned.

        Note that row and column are specified as one indexed, whereas xlsxwriter is
        zero indexed.
        """
        with ExcelWriteReadManager() as excel_manager:
            excel_manager.section_writer._write_supheader(
                worksheet=excel_manager.worksheet,
                row=row - 1,
                column=column - 1,
                num_columns=num_columns,
                supheader="Supheader",
                supheader_format={"bold": True},
            )
        return excel_manager.load_worksheet()

    @pytest.mark.parametrize("row, column", [(1, 2), (2, 1), (5, 10)])
    def test_supheader_written_to_correct_position(self, row, column):
        worksheet = self._create_worksheet_with_section_writer_and_write_supheader(row, column)  # fmt: skip
        assert worksheet.cell(row=row, column=column).value == "Supheader"

    def test_supheader_format_is_applied(self):
        worksheet = self._create_worksheet_with_section_writer_and_write_supheader()
        assert worksheet.cell(row=1, column=1).font.bold == True

    @pytest.mark.parametrize("num_columns", [2, 4, 10])
    def test_supheader_cell_is_merged(self, num_columns):
        worksheet = self._create_worksheet_with_section_writer_and_write_supheader(num_columns=num_columns)  # fmt: skip
        merged_cells = list(worksheet.merged_cells.ranges[0].cells)
        assert merged_cells == [(1, i) for i in range(1, num_columns + 1)]

    def test_no_merge_applied_when_written_with_only_one_column(self):
        worksheet = self._create_worksheet_with_section_writer_and_write_supheader(num_columns=1)  # fmt: skip
        assert len(worksheet.merged_cells.ranges) == 0
        assert worksheet.cell(row=1, column=1).value == "Supheader"


class TestTableSectionWriteSection:
    def _create_worksheet_with_section_writer_and_write_section(
        self, table_section, write_supheader
    ):
        """Creates an xlsxwriter.workbook, an xlsxwriter.worksheet and a
        TableSectionWriter. Then writes the `table_section` to the worksheet by using
        the TableSectionWriter._write_section method. The worksheet is then safed to a
        buffer and loaded with openpyxl. The loaded worksheet is returned.
        """
        with ExcelWriteReadManager() as excel_manager:
            excel_manager.section_writer._write_section(
                excel_manager.worksheet,
                table_section,
                start_row=0,
                start_column=0,
                write_supheader=write_supheader,
            )
        return excel_manager.load_worksheet()

    @pytest.mark.parametrize("write_supheader", [True, False])
    def test_correct_number_of_columns_written(self, table_section, write_supheader):
        worksheet = self._create_worksheet_with_section_writer_and_write_section(
            table_section, write_supheader=write_supheader
        )
        assert len(list(worksheet.columns)) == table_section.data.shape[1]

    @pytest.mark.parametrize("write_supheader, num_headers", [(True, 2), (False, 1)])
    def test_correct_number_of_rows_written(self, table_section, write_supheader, num_headers):  # fmt: skip
        worksheet = self._create_worksheet_with_section_writer_and_write_section(
            table_section, write_supheader=write_supheader
        )
        for column in worksheet.columns:
            assert len(column) == table_section.data.shape[0] + num_headers

    @pytest.mark.parametrize("write_supheader", [True, False])
    def test_supheader_correctly_added_or_ommitted(self, table_section, write_supheader):  # fmt: skip
        worksheet = self._create_worksheet_with_section_writer_and_write_section(
            table_section, write_supheader=write_supheader
        )
        # Note that xlsxwriter is zero indexed whereas pyopenxl is one indexed.
        if write_supheader:
            assert worksheet.cell(row=1, column=1).value == table_section.supheader
        else:
            assert worksheet.cell(row=1, column=1).value != table_section.supheader


class TestTableSectionWriterGetXlsxFormat:
    @pytest.fixture(autouse=True)
    def _init_section_writer(self):
        self.writer = writer.TableSectionWriter(Workbook())
        self.format_description = {"bold": True, "font_color": "#FF0000"}

    def test_xlsx_format_with_correct_properties_is_returned(self):
        xlsx_format = self.writer.get_xlsx_format(self.format_description)
        assert xlsx_format.bold == True and xlsx_format.font_color == "#FF0000"

    def test_multiple_calls_return_same_object(self):
        xlsx_format_1 = self.writer.get_xlsx_format(self.format_description)
        xlsx_format_2 = self.writer.get_xlsx_format(self.format_description)
        assert xlsx_format_1 is xlsx_format_2


class TestHashableFromDictionary:
    @pytest.mark.parametrize(
        "dictionary", [{"A": 1, "C": 2, "B": 1}, {"B": 1, "A": 1, "C": 2}]
    )
    def test_different_key_order_creates_same_hashable(self, dictionary):
        expected_hash = writer._hashable_from_dict({"A": 1, "B": 1, "C": 2})
        assert writer._hashable_from_dict(dictionary) == expected_hash

    def test_different_dict_values_create_different_hashable(self):
        hashable_1 = writer._hashable_from_dict({"A": 1})
        hashable_2 = writer._hashable_from_dict({"A": 2})
        assert hashable_1 != hashable_2
