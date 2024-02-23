import pytest
from unittest.mock import patch, MagicMock
import pathlib

import xlsxreport.appdir as appdir


@patch("xlsxreport.appdir.platformdirs.user_data_dir")
def test_locate_appdir(user_data_dir):
    appdir.locate_appdir()
    user_data_dir.assert_called_with(appname=appdir.APPNAME, appauthor=False)


def test_get_default_template_files_returns_yaml_files():
    default_template_files = list(appdir._get_default_template_files())
    assert len(default_template_files) > 0
    assert any([fn.name.endswith(".yaml") for fn in default_template_files])


# fmt: off
@pytest.fixture
def mock_setup():
    with patch("xlsxreport.appdir._get_default_template_files") as mock_get_files, \
         patch("pathlib.Path.is_file") as mock_is_file, \
         patch("shutil.copyfile") as mock_copyfile:
        mock_file = MagicMock(spec=pathlib.Path)
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
            self.mock_get_files.return_value[0], pathlib.Path("mock_dir", "file1")
        )

    def test_existing_files_are_not_copied(self, mock_setup):
        self.mock_is_file.return_value = True
        appdir._copy_default_templates("mock_dir", False)
        self.mock_copyfile.assert_not_called()

    def test_existing_files_are_overwritten_if_specified(self, mock_setup):
        self.mock_is_file.return_value = False
        appdir._copy_default_templates("mock_dir", True)
        self.mock_copyfile.assert_called_once_with(
            self.mock_get_files.return_value[0], pathlib.Path("mock_dir", "file1")
        )


class TestGetTemplatePath:
    @patch("pathlib.Path.is_file")
    def test_valid_filepath(self, mock_isfile):
        mock_isfile.return_value = True
        assert appdir.get_template_path("valid_filepath") == "valid_filepath"

    @patch("pathlib.Path.is_file")
    @patch("xlsxreport.appdir.get_appdir_templates")
    @patch("xlsxreport.appdir.locate_appdir")
    def test_invalid_filepath_found_in_appdir(
        self, mock_locate_appdir, mock_get_appdir_templates, mock_isfile
    ):
        mock_isfile.return_value = False
        mock_get_appdir_templates.return_value = ["file1", "file2"]
        mock_locate_appdir.return_value = "/path/to/appdir"
        assert appdir.get_template_path("file2") == "/path/to/appdir/file2"

    @patch("pathlib.Path.is_file")
    @patch("xlsxreport.appdir.get_appdir_templates")
    def test_invalid_filepath_not_found_in_appdir(
        self, mock_get_appdir_templates, mock_isfile
    ):
        mock_isfile.return_value = False
        mock_get_appdir_templates.return_value = []
        assert appdir.get_template_path("non_existing_file") is None
