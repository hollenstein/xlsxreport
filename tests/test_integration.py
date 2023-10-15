import os

import pytest
import pandas as pd
import openpyxl
import xlsxreport


@pytest.fixture()
def temp_excel_path(request, tmp_path):
    output_path = os.path.join(tmp_path, "output.xlsx")

    def teardown():
        if os.path.isfile(output_path):
            os.remove(output_path)

    request.addfinalizer(teardown)
    return output_path


def test_correct_creation_of_formatted_excel_file(temp_excel_path):
    config_path = "./tests/testdata/mq_protein_config.yaml"
    mq_path = "./tests/testdata/mq_proteinGroups.txt"
    table = pd.read_csv(mq_path, sep="\t")
    with xlsxreport.Reportbook(temp_excel_path) as reportbook:
        protein_sheet = reportbook.add_datasheet("Proteins")
        protein_sheet.apply_configuration(config_path)
        protein_sheet.add_data(table)
        protein_sheet.write_data()

    reference_excel_path = "./tests/testdata/mq_proteinGroups.xlsx"
    with open(temp_excel_path, "rb") as f1, open(reference_excel_path, "rb") as f2:
        wb1 = openpyxl.load_workbook(f1)
        wb2 = openpyxl.load_workbook(f2)

        # Compare the number of sheets
        assert len(wb1.sheetnames) == len(wb2.sheetnames)

        # Compare the sheet names
        for sheet1, sheet2 in zip(wb1, wb2):
            assert sheet1.title == sheet2.title

        # Compare the cell values and formatting
        for sheet1, sheet2 in zip(wb1, wb2):
            for row1, row2 in zip(sheet1.iter_rows(), sheet2.iter_rows()):
                for cell1, cell2 in zip(row1, row2):
                    assert cell1.value == cell2.value
                    assert cell1._style == cell2._style
                    assert str(cell1.number_format) == str(cell2.number_format)
                    assert str(cell1.font) == str(cell2.font)
                    assert str(cell1.border) == str(cell2.border)
                    assert str(cell1.fill) == str(cell2.fill)
                    assert str(cell1.alignment) == str(cell2.alignment)

        # Compare the conditional formatting rules
        for sheet1, sheet2 in zip(wb1, wb2):
            for rule1, rule2 in zip(
                sheet1.conditional_formatting._cf_rules,
                sheet2.conditional_formatting._cf_rules,
            ):
                assert rule1.cells == rule2.cells
                for cfrule1, cfrul2 in zip(rule1.rules, rule2.rules):
                    assert str(cfrule1) == str(cfrul2)
