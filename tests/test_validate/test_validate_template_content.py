import pytest

import xlsxreport.validate as validate


def test_errors_from_expected_and_unexpected_main_section_entries():
    report_template = {e.value: {} for e in validate.MainSections}
    del report_template["formats"]
    report_template["NOT A MAIN SECTION"] = {}
    errors = validate.validate_template_content(report_template)

    expected_error_entries = [
        (("formats",), validate.ValidationErrorType.MISSING_PARAMETER),
        (("NOT A MAIN SECTION",), validate.ValidationErrorType.UNKNOWN_PARAMTER),
    ]
    for error_entry in expected_error_entries:
        assert error_entry in [(err.field, err.error_type) for err in errors]


def test_errors_from_invalid_format_and_conditional_format_descriptions():
    report_template = {
        "formats": {"format": "NOT VALID"},
        "conditional_formats": {"format": "NOT VALID"},
    }
    errors = validate.validate_template_content(report_template)

    expected_error_entries = [
        (("formats", "format"), validate.ValidationErrorType.INVALID_FORMAT),
        (("conditional_formats", "format"), validate.ValidationErrorType.INVALID_FORMAT)  # fmt: skip
    ]
    for error_entry in expected_error_entries:
        assert error_entry in [(err.field, err.error_type) for err in errors]


def test_errors_from_expected_and_unexpected_settings_entries():
    report_template = {"settings": {s: {} for s in validate.SETTINGS_SCHEMA}}
    del report_template["settings"]["header_height"]
    report_template["settings"]["INVALID"] = ""
    errors = validate.validate_template_content(report_template)

    expected_error_entries = [
        (("settings", "header_height"), validate.ValidationErrorType.MISSING_PARAMETER),
        (("settings", "INVALID"), validate.ValidationErrorType.UNKNOWN_PARAMTER),
    ]
    for error_entry in expected_error_entries:
        assert error_entry in [(err.field, err.error_type) for err in errors]


def test_errors_from_unused_and_undefined_formats():
    report_template = {
        "sections": {"section": {"format": "UNDEFINED"}},
        "formats": {"UNUSED": {}},
    }
    errors = validate.validate_template_content(report_template)

    expected_error_entries = [
        (("formats", "UNDEFINED"), validate.ValidationErrorType.UNDEFINED_FORMAT),
        (("formats", "header"), validate.ValidationErrorType.UNDEFINED_FORMAT),
        (("formats", "UNUSED"), validate.ValidationErrorType.UNUSED_FORMAT),
    ]
    for error_entry in expected_error_entries:
        assert error_entry in [(err.field, err.error_type) for err in errors]


def test_errors_from_unused_and_undefined_conditional_formats():
    report_template = {
        "sections": {"section": {"conditional_format": "UNDEFINED"}},
        "conditional_formats": {"UNUSED": {}},
    }
    errors = validate.validate_template_content(report_template)

    expected_error_entries = [
        (("conditional_formats", "UNDEFINED"), validate.ValidationErrorType.UNDEFINED_FORMAT),  # fmt: skip
        (("conditional_formats", "UNUSED"), validate.ValidationErrorType.UNUSED_FORMAT),  # fmt: skip
    ]
    for error_entry in expected_error_entries:
        assert error_entry in [(err.field, err.error_type) for err in errors]
