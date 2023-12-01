import pytest

import xlsxreport.validate as validate


class TestValidateExpectedMainSections:
    def test_no_error_when_all_main_sections_are_present(self):
        report_template = {e.value: {} for e in validate.MainSections}
        errors = validate.validate_expected_main_sections(report_template)
        assert len(errors) == 0

    def test_error_when_a_main_section_is_missing(self):
        report_template = {}
        errors = validate.validate_expected_main_sections(report_template)
        assert len(errors) == 4
        assert all([err.error_type == validate.ValidationErrorType.MISSING_PARAMETER for err in errors])  # fmt: skip
        assert all([err.error_level == validate.ErrorLevel.WARNING for err in errors])
        assert sorted([err.field for err in errors]) == sorted([(e.value,) for e in validate.MainSections])  # fmt: skip


class TestValidateUnexpectedMainSections:
    def test_no_error_when_expected_main_sections_are_present(self):
        report_template = {e.value: {} for e in validate.MainSections}
        errors = validate.validate_unexpected_main_sections(report_template)
        assert len(errors) == 0

    def test_errors_for_unexpected_main_section_names(self):
        report_template = {"UNEXPECTED": "a", "NOT EXPECTED": "b"}
        errors = validate.validate_unexpected_main_sections(report_template)
        assert len(errors) == 2
        assert all([err.error_type == validate.ValidationErrorType.UNKNOWN_PARAMTER for err in errors])  # fmt: skip
        assert all([err.error_level == validate.ErrorLevel.WARNING for err in errors])
        assert sorted([err.field for err in errors]) == sorted([("UNEXPECTED",), ("NOT EXPECTED",)])  # fmt: skip
