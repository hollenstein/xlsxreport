#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
XlsxReport is a Python package to automatically generate formatted excel reports
from quantitative mass spectrometry result tables. YAML config files are used to
describe the content of a result table and the format of the excel report.
"""


from .writer import Reportbook
from .appdir import locate_data_dir, get_config_file, setup_data_dir


__author__ = "David M. Hollenstein"
__license__ = "Apache 2.0"
__version__ = "0.0.7"
