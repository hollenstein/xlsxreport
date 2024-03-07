import pytest
import xlsxwriter
import warnings

import xlsxreport.validate as validate


class TestValidateConditionalFormatDescriptions:
    @pytest.fixture(autouse=True)
    def _init_formats(self):
        self.valid_formats = {
            "test": {
                "type": "2_color_scale",
                "min_type": "num",
                "min_value": 0,
                "min_color": "#ffffbf",
                "max_type": "percentile",
                "max_value": 99.9,
                "max_color": "#f25540",
            }
        }
        self.invalid_formats = {
            "test_1": {"invalid format key": 2},
            "test_2": {"invalid format key": 2},
        }

    def test_no_errors_with_valid_format(self):
        errors = validate.validate_conditional_format_descriptions(self.valid_formats)
        assert len(errors) == 0

    def test_invalid_format_is_reported(self):
        errors = validate.validate_conditional_format_descriptions(self.invalid_formats)
        assert len(errors) == len(self.invalid_formats)

    def test_errors_have_the_correct_error_type(self):
        errors = validate.validate_conditional_format_descriptions(self.invalid_formats)
        assert all([e.error_type == validate.ValidationErrorType.INVALID_FORMAT for e in errors])  # fmt: skip

    def test_errors_have_the_correct_error_level(self):
        errors = validate.validate_conditional_format_descriptions(self.invalid_formats)
        assert all([e.error_level == validate.ErrorLevel.ERROR for e in errors])


class TestValidateFormatDescriptions:
    @pytest.fixture(autouse=True)
    def _init_formats(self):
        self.valid_formats = {
            "str": {"align": "left", "num_format": "0"},
            "header": {"bold": True, "align": "center", "valign": "vcenter"},
        }

        self.invalid_formats = {
            "str": {"invalid format key": 2},
            "int": {"invalid format key": 2},
        }

    def test_no_errors_with_valid_format(self):
        errors = validate.validate_format_descriptions(self.valid_formats)
        assert len(errors) == 0

    def test_invalid_format_is_reported(self):
        errors = validate.validate_format_descriptions(self.invalid_formats)
        assert len(errors) == len(self.invalid_formats)

    def test_errors_have_the_correct_error_type(self):
        errors = validate.validate_format_descriptions(self.invalid_formats)
        assert all([e.error_type == validate.ValidationErrorType.INVALID_FORMAT for e in errors])  # fmt: skip

    def test_errors_have_the_correct_error_level(self):
        errors = validate.validate_format_descriptions(self.invalid_formats)
        assert all([e.error_level == validate.ErrorLevel.ERROR for e in errors])


class TestValidateUnusedFormats:
    def test_correct_error_when_format_is_unused(self):
        template_document = {"formats": {"f1": {}}}
        errors = validate.validate_unused_formats(template_document)
        assert errors[0].error_type == validate.ValidationErrorType.UNUSED_FORMAT
        assert errors[0].error_level == validate.ErrorLevel.INFO
        assert errors[0].field == (validate.MainSections.FORMATS.value, "f1")

    def test_all_formats_create_errors_when_sections_section_is_absent(self):
        template_document = {"formats": {"f1": {}, "f2": {}, "f3": {}}}
        errors = validate.validate_unused_formats(template_document)
        assert len(errors) == 3

    def test_only_unused_formats_create_errors(self):
        template_document = {
            "sections": {"s1": {"format": "USED FORMAT"}},
            "formats": {"USED FORMAT": {}, "UNUSED FORMAT": {}},
        }
        errors = validate.validate_unused_formats(template_document)
        assert len(errors) == 1
        assert "UNUSED FORMAT" in errors[0].field
        assert "USED FORMAT" not in errors[0].field

    def test_no_errors_when_special_formats_are_unused(self):
        template_document = {"formats": {f: {} for f in validate.SPECIAL_FORMATS}}
        errors = validate.validate_unused_formats(template_document)
        assert len(errors) == 0

    def test_no_errors_when_format_section_is_absent(self):
        template_document = {"sections": {"s1": {"format": "UNUSED FORMAT"}}}
        errors = validate.validate_unused_formats(template_document)
        assert len(errors) == 0


class TestValidateUndefinedFormats:
    def test_correct_error_when_format_is_undefined(self):
        template_document = {"sections": {"s1": {"format": "UNUSED FORMAT"}}}
        errors = validate.validate_undefined_formats(template_document)
        assert errors[0].error_type == validate.ValidationErrorType.UNDEFINED_FORMAT
        assert errors[0].error_level == validate.ErrorLevel.ERROR
        assert errors[0].field == (validate.MainSections.FORMATS.value, "UNUSED FORMAT")

    def test_only_undefined_formats_create_errors(self):
        template_document = {
            "sections": {
                "s1": {"format": "UNDEFINED FORMAT 1"},
                "s2": {"format": "UNDEFINED FORMAT 2"},
                "s3": {"format": "DEFINED FORMAT"},
            },
            "formats": {"DEFINED FORMAT": {}},
        }
        errors = validate.validate_undefined_formats(template_document)
        assert len(errors) == 2
        assert "DEFINED FORMAT" not in errors[0].field

    def test_no_errors_when_sections_section_is_absent(self):
        template_document = {"formats": {"f1": {}, "f2": {}}}
        errors = validate.validate_undefined_formats(template_document)
        assert len(errors) == 0


class TestValidateSpecialFormatsDefined:
    def test_correct_error_when_special_format_is_undefined(self):
        template_document = {"formats": {f: {} for f in validate.SPECIAL_FORMATS[1:]}}
        errors = validate.validate_special_formats_defined(template_document)
        assert errors[0].error_type == validate.ValidationErrorType.UNDEFINED_FORMAT
        assert errors[0].error_level == validate.ErrorLevel.INFO
        assert errors[0].field == (validate.MainSections.FORMATS.value, validate.SPECIAL_FORMATS[0])  # fmt: skip

    def test_no_errors_when_all_special_formats_are_defined(self):
        template_document = {"formats": {f: {} for f in validate.SPECIAL_FORMATS}}
        errors = validate.validate_special_formats_defined(template_document)
        assert len(errors) == 0

    def test_each_absent_special_format_creates_an_error(self):
        template_document = {"formats": {}}
        errors = validate.validate_special_formats_defined(template_document)
        assert len(errors) == len(validate.SPECIAL_FORMATS)


class TestValidateUnusedConditionalFormats:
    def test_correct_error_when_conditional_format_is_unused(self):
        template_document = {"conditional_formats": {"f1": {}}}
        errors = validate.validate_unused_conditional_formats(template_document)
        assert errors[0].error_type == validate.ValidationErrorType.UNUSED_FORMAT
        assert errors[0].error_level == validate.ErrorLevel.INFO
        assert errors[0].field == (validate.MainSections.CONDITIONAL_FORMATS.value, "f1")  # fmt: skip

    def test_all_formats_create_errors_when_sections_section_is_absent(self):
        template_document = {"conditional_formats": {"f1": {}, "f2": {}, "f3": {}}}
        errors = validate.validate_unused_conditional_formats(template_document)
        assert len(errors) == 3

    def test_only_unused_formats_create_errors(self):
        template_document = {
            "sections": {"s1": {"conditional_format": "USED FORMAT"}},
            "conditional_formats": {"USED FORMAT": {}, "UNUSED FORMAT": {}},
        }
        errors = validate.validate_unused_conditional_formats(template_document)
        assert len(errors) == 1
        assert "UNUSED FORMAT" in errors[0].field
        assert "USED FORMAT" not in errors[0].field

    def test_no_errors_when_format_section_is_absent(self):
        template_document = {"sections": {"s1": {"conditional_format": "UNUSED FORMAT"}}}  # fmt: skip
        errors = validate.validate_unused_conditional_formats(template_document)
        assert len(errors) == 0


class TestValidateUndefinedConditionalFormats:
    def test_correct_error_when_format_is_undefined(self):
        template_document = {"sections": {"s1": {"conditional_format": "UNUSED FORMAT"}}}  # fmt: skip
        errors = validate.validate_undefined_conditional_formats(template_document)
        assert errors[0].error_type == validate.ValidationErrorType.UNDEFINED_FORMAT
        assert errors[0].error_level == validate.ErrorLevel.ERROR
        assert errors[0].field == (validate.MainSections.CONDITIONAL_FORMATS.value, "UNUSED FORMAT")  # fmt: skip

    def test_only_undefined_formats_create_errors(self):
        template_document = {
            "sections": {
                "s1": {"conditional_format": "UNDEFINED FORMAT 1"},
                "s2": {"conditional_format": "UNDEFINED FORMAT2"},
                "s3": {"conditional_format": "DEFINED FORMAT"},
            },
            "conditional_formats": {"DEFINED FORMAT": {}},
        }
        errors = validate.validate_undefined_conditional_formats(template_document)
        assert len(errors) == 2
        assert "DEFINED FORMAT" not in errors[0].field

    def test_no_errors_when_sections_section_is_absent(self):
        template_document = {"conditional_formats": {"f1": {}, "f2": {}}}
        errors = validate.validate_undefined_conditional_formats(template_document)
        assert len(errors) == 0
