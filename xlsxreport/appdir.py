import errno
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
    templates = []
    templates.extend([fn.name for fn in pathlib.Path(locate_appdir()).glob(f"*.yaml")])
    templates.extend([fn.name for fn in pathlib.Path(locate_appdir()).glob(f"*.yml")])
    return templates


def setup_appdir(overwrite_templates: bool = False) -> None:
    """Creates a user specific app data directory and copies default template files.

    Args:
        overwrite_templates: If True, existing template files will be overwritten.
    """
    data_dir = locate_appdir()
    if not os.path.isdir(data_dir):
        os.makedirs(data_dir)
    _copy_default_templates(data_dir, overwrite_templates)


def get_template_path(template: str) -> str:
    """Returns the path to the specified template file.

    Args:
        template: The path to the template file or the name of the template file to
            locate in the app directory. If a valid filepath to an existing file is
            provided, the specified path is returned without checking if the file exists
            also exists in the app directory.

    Returns:
        The path to the specified template file.

    Raises:
        FileNotFoundError: If the specified template file does not exist.
    """
    if pathlib.Path(template).is_file():
        return template
    elif template in get_appdir_templates():
        return pathlib.Path(locate_appdir(), template).as_posix()
    else:
        raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), template)


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
