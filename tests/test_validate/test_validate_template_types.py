import pytest

import xlsxreport.validate as validate


class TestValidateDocumentEntryTypes_MainSectionTypes:
    @pytest.fixture(autouse=True)
    def _init_template_document(self):
        self.document = {key: {} for key in validate.TEMPLATE_SCHEMA}

    def test_valid_main_section_types_do_not_create_errors(self):
        errors = validate.validate_document_entry_types(self.document)
        assert len(errors) == 0

    def test_missing_main_sections_do_not_create_errors(self):
        errors = validate.validate_document_entry_types({})
        assert len(errors) == 0

    def test_invalid_main_section_names_are_not_checked(self):
        document = {"NOT A MAIN SECTION 1": {}, "NOT A MAIN SECTION 2": "A STRING"}
        errors = validate.validate_document_entry_types(document)
        assert len(errors) == 0

    def test_invalid_main_section_types_are_reported(self):
        self.document = {key: "NOT A DICT" for key in self.document}
        errors = validate.validate_document_entry_types(self.document)
        assert len(errors) == len(self.document)
        assert errors[0].error_type == validate.ValidationErrorType.TYPE_ERROR
        assert errors[0].error_level == validate.ErrorLevel.CRITICAL


class TestValidateDocumentEntryTypes_SectionsParameters:
    def test_invalid_sections_parameter_types_are_reported(self):
        template_document = {
            "sections": {
                "Invalid section": "NOT A DICT",
                "Valid section": {"format": 123, 123: "A NUMBER AS KEY"},
                123: {"format": 345},
            }
        }
        errors = validate.validate_document_entry_types(template_document)
        error_fields = [err.field for err in errors]
        expected_fields = [
            ("sections", 123),
            ("sections", 123, "format"),
            ("sections", "Invalid section"),
            ("sections", "Valid section", "format"),
        ]
        assert all([field in error_fields for field in expected_fields])
        assert len(errors) == len(expected_fields)

    def test_invalid_section_parameter_names_are_ignored(self):
        template_document = {"sections": {"Invalid section": {"NOT A PARAM": 1}}}
        errors = validate.validate_document_entry_types(template_document)
        assert len(errors) == 0


class TestValidateDocumentEntryTypes_FormatsParameters:
    def test_formats_entry_that_is_not_a_dict_creates_error(self):
        template_document = {"formats": {"Invalid format": "NOT A DICT"}}
        errors = validate.validate_document_entry_types(template_document)
        assert len(errors) == 1
        assert errors[0].field == ("formats", "Invalid format")
        assert errors[0].error_type == validate.ValidationErrorType.TYPE_ERROR
        assert errors[0].error_level == validate.ErrorLevel.CRITICAL

    def test_formats_entries_of_type_dict_creat_no_errors(self):
        template_document = {"formats": {f: {} for f in ["a", "b", "c"]}}
        errors = validate.validate_document_entry_types(template_document)
        assert len(errors) == 0


class TestValidateDocumentEntryTypes_ConditionalFormatsParameters:
    def test_formats_entry_that_is_not_a_dict_creates_error(self):
        template_document = {"conditional_formats": {"Invalid format": "NOT A DICT"}}
        errors = validate.validate_document_entry_types(template_document)
        assert len(errors) == 1
        assert errors[0].field == ("conditional_formats", "Invalid format")
        assert errors[0].error_type == validate.ValidationErrorType.TYPE_ERROR
        assert errors[0].error_level == validate.ErrorLevel.CRITICAL

    def test_formats_entries_of_type_dict_creat_no_errors(self):
        template_document = {"conditional_formats": {f: {} for f in ["a", "b", "c"]}}
        errors = validate.validate_document_entry_types(template_document)
        assert len(errors) == 0


class TestValidateDocumentEntryTypes_SettingsParameters:
    @pytest.fixture(autouse=True)
    def _init_default_settings(self):
        self.settings = {key: info["default"] for key, info in validate.SETTINGS_SCHEMA.items()}  # fmt: skip

    def test_invalid_settings_parameter_types_are_reported(self):
        template_document = {
            "settings": {
                "column_width": "",
                "header_height": -1,
                "log2_tag": 0,
                "write_supheader": 1,
            }
        }
        errors = validate.validate_document_entry_types(template_document)
        assert len(errors) == len(template_document["settings"])

    def test_no_errors_when_using_valid_parameter_types(self):
        template_document = {"settings": self.settings}
        errors = validate.validate_document_entry_types(template_document)
        assert len(errors) == 0

    def test_errors_have_the_correct_error_type(self):
        template_document = {"settings": {k: -1 for k in self.settings}}
        errors = validate.validate_document_entry_types(template_document)
        assert errors
        assert all([e.error_type == validate.ValidationErrorType.TYPE_ERROR for e in errors])  # fmt: skip

    def test_errors_have_the_correct_error_level(self):
        template_document = {"settings": {k: -1 for k in self.settings}}
        errors = validate.validate_document_entry_types(template_document)
        assert errors
        assert all([e.error_level == validate.ErrorLevel.CRITICAL for e in errors])


class TestFlatteCerberusErrors:
    def test_with_nested_errors(self):
        cerberus_errors = {
            "formats": ["must be of dict type"],
            "sections": [
                {
                    123: [
                        "must be of string type",
                        {"format": ["must be of string type"]},
                    ],
                    "Invalid section": ["must be of dict type"],
                    "Valid section": [{"format": ["must be of string type"]}],
                }
            ],
            "settings": [{"header_height": ["must be of float type"]}],
        }

        flat_errors = [
            (("formats",), "must be of dict type"),
            (("sections", 123), "must be of string type"),
            (("sections", 123, "format"), "must be of string type"),
            (("sections", "Invalid section"), "must be of dict type"),
            (("sections", "Valid section", "format"), "must be of string type"),
            (("settings", "header_height"), "must be of float type"),
        ]

        assert validate._flatten_cerberus_errors(cerberus_errors) == flat_errors

    def test_when_no_errors(self):
        cerberus_errors = {}
        assert validate._flatten_cerberus_errors(cerberus_errors) == []
