""" Test docstring of xlsxreport.appdir module


The module contains the following functions:

- `locate_data_dir()` - Returns the XlsxReport user data directory.
- `setup_data_dir()` - Creates a XlsxReport user data directory.
- `get_config_file(filename)` - Returns the full path of a config file in the data dir.
"""

import os
import shutil
from typing import Union

import appdirs


def locate_data_dir():
    APPNAME = "XlsxReport"
    return appdirs.user_data_dir(appname=APPNAME, appauthor=False)


def setup_data_dir():
    """Creates a user specific app data directory and copies default config files."""
    data_dir = locate_data_dir()
    if not os.path.isdir(data_dir):
        os.makedirs(data_dir)

    config_dir = os.path.join(os.path.dirname(__file__), "default_config")
    for filename in os.listdir(config_dir):
        src_path = os.path.join(config_dir, filename)
        dest_path = os.path.join(data_dir, filename)
        shutil.copy(src_path, dest_path)


def get_config_file(filename: str) -> Union[str, None]:
    """Returns the file path if filename is present in the app data directory.

    Args:
        filename: Config filename

    Returns:
        Full path to config file or None
    """
    file_path = None
    data_dir = locate_data_dir()

    for subdir, dirs, files in os.walk(data_dir):
        for file in files:
            if file == filename:
                file_path = os.path.join(subdir, file)
    return file_path
