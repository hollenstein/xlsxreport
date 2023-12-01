import pytest

from xlsxreport.validate import SETTINGS_SCHEMA
import pytest
import xlsxreport.validate as validate


@pytest.fixture
def default_settings():
    return {key: info["default"] for key, info in SETTINGS_SCHEMA.items()}


class TestValidateUnexpectedSettingsParameters:
    def test_unknown_section_parameters_are_reported(self):
        settings = {"unknown key1": 1, "unknown key2": 2}
        errors = validate.validate_unexpected_settings_parameters(settings)
        assert len(errors) == 2

    def test_no_errors_when_using_all_existing_parameters(self, default_settings):
        errors = validate.validate_unexpected_settings_parameters(default_settings)
        assert len(errors) == 0

    def test_errors_have_the_correct_error_type(self):
        errors = validate.validate_unexpected_settings_parameters({"unknown key1": 1})
        assert all([e.error_type == validate.ValidationErrorType.UNKNOWN_PARAMTER for e in errors])  # fmt: skip

    def test_errors_have_the_correct_error_level(self):
        errors = validate.validate_unexpected_settings_parameters({"unknown key1": 1})
        assert all([e.error_level == validate.ErrorLevel.WARNING for e in errors])


class TestValidateExpectedSettingsParameters:
    def test_missing_section_parameters_are_reported(self, default_settings):
        errors = validate.validate_expected_settings_parameters({})
        assert len(errors) == len(default_settings)

    def test_no_errors_when_all_parameters_are_present(self, default_settings):
        errors = validate.validate_expected_settings_parameters(default_settings)
        assert len(errors) == 0

    def test_errors_have_the_correct_error_type(self):
        errors = validate.validate_expected_settings_parameters({})
        assert errors
        assert all([e.error_type == validate.ValidationErrorType.MISSING_PARAMETER for e in errors])  # fmt: skip

    def test_errors_have_the_correct_error_level(self):
        errors = validate.validate_expected_settings_parameters({})
        assert errors
        assert all([e.error_level == validate.ErrorLevel.INFO for e in errors])
