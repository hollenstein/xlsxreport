"""
XlsxReport is a Python package to automatically generate formatted excel reports
from quantitative mass spectrometry result tables. YAML template files are used to
describe how the content of a result table should be formatted in the Excel report.
"""


from .writer import Reportbook
from .appdir import get_template_path


__author__ = "David M. Hollenstein"
__license__ = "Apache 2.0"
__version__ = "0.0.8"
