import os
import shutil
from typing import Union

import appdirs


def locate_data_dir():
    APPNAME = "XlsxReport"
    return appdirs.user_data_dir(appname=APPNAME, appauthor=False)


def setup_data_dir():
    """Creates a user specific app data directory and copies default template files."""
    data_dir = locate_data_dir()
    if not os.path.isdir(data_dir):
        os.makedirs(data_dir)

    template_dir = os.path.join(os.path.dirname(__file__), "default_templates")
    for filename in os.listdir(template_dir):
        src_path = os.path.join(template_dir, filename)
        dest_path = os.path.join(data_dir, filename)
        shutil.copy(src_path, dest_path)


def get_template_file(filename: str) -> Union[str, None]:
    """Returns the file path if filename is present in the app data directory."""
    file_path = None
    data_dir = locate_data_dir()

    for subdir, dirs, files in os.walk(data_dir):
        for file in files:
            if file == filename:
                file_path = os.path.join(subdir, file)
    return file_path
