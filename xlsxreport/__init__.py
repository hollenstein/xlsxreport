"""
XlsxReport is a Python package to automatically generate formatted excel reports
from quantitative mass spectrometry result tables. YAML template files are used to
describe how the content of a table should be formatted in the Excel report.

Exposes the following functions and classes:
    - get_template_path (function): Returns the path to a template file. Returns the
        specified path if the file exists, otherwise looks for the file in the
        XlsxReport app directory.
    - prepare_compiled_sections (function): Generate a list of `CompiledSection`s from a
        table (pandas.DataFrame) and a `TableTemplate`.
    - TableTemplate (class): Python representation of a YAML template file. Can be
        used to load, edit, and save table template files.
    - ReportBuilder (class): Provides an interface for building multi-tab Excel reports.
    - SectionWriter (class): Provides an interface for writing a list of
        `CompiledSection`s to an Excel file by using the `xlsxwriter` package.
"""

from xlsxreport.appdir import get_template_path
from xlsxreport.builder import ReportBuilder
from xlsxreport.compiler import prepare_compiled_sections
from xlsxreport.template import TableTemplate
from xlsxreport.writer import SectionWriter

__author__ = "David M. Hollenstein"
__license__ = "Apache 2.0"
__version__ = "0.1.1"
