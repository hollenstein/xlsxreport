from unittest.mock import patch
import os

import xlsxreport.appdir as appdir


@patch("xlsxreport.appdir.platformdirs.user_data_dir")
def test_locate_appdir(user_data_dir):
    appdir.locate_appdir()
    user_data_dir.assert_called_with(appname=appdir.APPNAME, appauthor=False)


def test_get_default_template_files_returns_yaml_files():
    default_template_files = appdir._get_default_template_files()
    assert any([fn.name.endswith(".yaml") for fn in default_template_files])


import pytest
from unittest.mock import patch, MagicMock
from pathlib import Path


class TestCopyDefaultTemplates:
    @patch("xlsxreport.appdir._get_default_template_files")
    @patch("pathlib.Path.is_file")
    @patch("shutil.copyfile")
    def test_files_are_copied(self, mock_copyfile, mock_is_file, mock_get_files):
        mock_file = MagicMock(spec=Path)
        mock_file.name = "file1"
        mock_get_files.return_value = [mock_file]
        mock_is_file.return_value = False

        appdir._copy_default_templates("mock_dir", False)
        mock_copyfile.assert_called_once_with(mock_file, Path("mock_dir", "file1"))

    @patch("xlsxreport.appdir._get_default_template_files")
    @patch("pathlib.Path.is_file")
    @patch("shutil.copyfile")
    def test_existing_files_are_not_copied(self, mock_copyfile, mock_is_file, mock_get_files):  # fmt: skip
        mock_file = MagicMock(spec=Path)
        mock_file.name = "file1"
        mock_get_files.return_value = [mock_file]
        mock_is_file.return_value = True

        appdir._copy_default_templates("mock_dir", False)
        mock_copyfile.assert_not_called()

    @patch("xlsxreport.appdir._get_default_template_files")
    @patch("pathlib.Path.is_file")
    @patch("shutil.copyfile")
    def test_existing_files_are_overwritten_if_specified(self, mock_copyfile, mock_is_file, mock_get_files):  # fmt: skip
        mock_file = MagicMock(spec=Path)
        mock_file.name = "file1"
        mock_get_files.return_value = [mock_file]
        mock_is_file.return_value = True

        appdir._copy_default_templates("mock_dir", True)
        mock_copyfile.assert_called_once_with(mock_file, Path("mock_dir", "file1"))


# fmt: off
@pytest.fixture
def mock_setup():
    with patch("xlsxreport.appdir._get_default_template_files") as mock_get_files, \
         patch("pathlib.Path.is_file") as mock_is_file, \
         patch("shutil.copyfile") as mock_copyfile:
        mock_file = MagicMock(spec=Path)
        mock_file.name = "file1"
        mock_get_files.return_value = [mock_file]

        yield mock_get_files, mock_is_file, mock_copyfile
# fmt: on


class TestCopyDefaultTemplates:
    @pytest.fixture(autouse=True)
    def _init(self, mock_setup):
        self.mock_get_files, self.mock_is_file, self.mock_copyfile = mock_setup

    def test_files_are_copied(self, mock_setup):
        self.mock_is_file.return_value = False
        appdir._copy_default_templates("mock_dir", False)
        self.mock_copyfile.assert_called_once_with(
            self.mock_get_files.return_value[0], Path("mock_dir", "file1")
        )

    def test_existing_files_are_not_copied(self, mock_setup):
        self.mock_is_file.return_value = True
        appdir._copy_default_templates("mock_dir", False)
        self.mock_copyfile.assert_not_called()

    def test_existing_files_are_overwritten_if_specified(self, mock_setup):
        self.mock_is_file.return_value = False
        appdir._copy_default_templates("mock_dir", True)
        self.mock_copyfile.assert_called_once_with(
            self.mock_get_files.return_value[0], Path("mock_dir", "file1")
        )
