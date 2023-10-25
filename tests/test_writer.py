from io import BytesIO
import pandas as pd
import pytest
import openpyxl
import xlsxreport.excel_writer as writer
import xlsxreport.compiler as compiler
from xlsxwriter import Workbook as Workbook


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


class TestTableSectionWriteColumn:
    @pytest.fixture(autouse=True)
    def _init_section_writer_with_write_column(self):
        # Note that xlsxwriter is zero indexed whereas pyopenxl is one indexed
        row, column = 2, 2
        with ExcelWriteReadManager() as excel_manager:
            workbook = excel_manager.workbook
            worksheet = excel_manager.worksheet
            section_writer = writer.TableSectionWriter(workbook)
            section_writer._write_column(
                worksheet,
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
        self.header_cell = self.worksheet.cell(row=2, column=2)
        self.written_column = list(self.worksheet.columns)[1]
        self.written_cells = self.written_column[1:]

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
        assert self.header_cell.font.bold == True
        assert self.header_cell.border.bottom.style == "medium"

    def test_column_format_is_applied(self):
        column_cells = self.written_cells[1:]  # without header cell
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


class TestTableSectionWriteSection:
    def _create_worksheet_with_section_writer_and_write_section(
        self, table_section, write_supheader
    ):
        # Note that xlsxwriter is zero indexed whereas pyopenxl is one indexed
        row, column = 1, 1
        with ExcelWriteReadManager() as excel_manager:
            workbook = excel_manager.workbook
            worksheet = excel_manager.worksheet
            section_writer = writer.TableSectionWriter(workbook)
            section_writer._write_section(
                worksheet,
                table_section,
                start_row=row - 1,
                start_column=column - 1,
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
