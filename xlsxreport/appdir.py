import os
import pathlib
import shutil
from typing import Iterator, Union

import platformdirs


APPNAME = "XlsxReport"


def locate_appdir() -> str:
    """Returns the path to the user specific app data directory."""
    return platformdirs.user_data_dir(appname=APPNAME, appauthor=False)


def get_appdir_templates() -> list[str]:
    """Returns a list of template filenames located in the user app data directory."""
    return [fn.name for fn in pathlib.Path(locate_appdir()).glob("*.yaml")]


def setup_appdir(overwrite_templates: bool = False) -> None:
    """Creates a user specific app data directory and copies default template files.

    Args:
        overwrite_templates: If True, existing template files will be overwritten.
    """
    data_dir = locate_appdir()
    if not os.path.isdir(data_dir):
        os.makedirs(data_dir)
    _copy_default_templates(data_dir, overwrite_templates)


def get_template_path(template: str) -> Union[str, None]:
    """Returns the path to the specified template file.

    Args:
        template: The name of the template file to locate in the app directory or the
            path to the template file.

    Returns:
        The path to the specified template file or None if the file was not found.
    """
    if pathlib.Path.is_file(template):
        return template
    elif template in get_appdir_templates():
        return pathlib.Path(locate_appdir(), template).as_posix()
    else:
        return None


def _copy_default_templates(directory: str, overwrite: bool) -> None:
    """Copies default template files to the specified directory."""
    for source in _get_default_template_files():
        destiny = pathlib.Path(directory, source.name)
        if not pathlib.Path.is_file(destiny) or overwrite:
            shutil.copyfile(source, destiny)


def _get_default_template_files() -> Iterator[pathlib.Path]:
    """Returns a list of default template files included in the XlsxReport package."""
    default_template_dir = pathlib.Path(__file__).parent / "default_templates"
    return default_template_dir.iterdir()
