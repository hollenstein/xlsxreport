import os

import pandas as pd
import xlsxwriter

from xlsxreport.excel_writer import TableSectionWriter
from xlsxreport.template import ReportTemplate
from xlsxreport.compiler import compile_table_sections

root = "D:/python/xlsxreport"
config_path = "./tests/testdata/mq_protein_config.yaml"
mq_path = "./tests/testdata/mq_proteinGroups.txt"

config_path = os.path.join(root, config_path)
mq_path = os.path.join(root, mq_path)
excel_path = "D:/test_new_xlsxwriter.xlsx"

table = pd.read_csv(mq_path, sep="\t").fillna("")
report_template = ReportTemplate.load(config_path)
table_sections = compile_table_sections(report_template, table)

with xlsxwriter.Workbook(excel_path) as workbook:
    worksheet = workbook.add_worksheet("Proteins")
    section_writer = TableSectionWriter(workbook)
    section_writer.write_sections(worksheet, table_sections)
