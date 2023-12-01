import os

import pytest

import xlsxreport.validate as validate


def _create_yaml_file(tmp_path, content: str = ""):
    file_path = tmp_path / "test_file.yaml"

    with open(file_path, "w") as f:
        f.write(content)

    return file_path


class TestValidateTemplateFileIntegrity:
    @pytest.mark.parametrize(
        "content",
        ['"Parser error', "['not a', 'dictrionary']"],
    )
    def test_error_with_invalid_yaml_file_content(self, content, tmp_path):
        yaml_path = _create_yaml_file(tmp_path, content)
        errors = validate.validate_template_file_integrity(yaml_path)
        assert len(errors) == 1
        assert errors[0].error_level == validate.ErrorLevel.CRITICAL

    def test_no_error_with_valid_yaml_file_content(self, tmp_path):
        content = "Key1: Value1\nKey2: Value2"
        yaml_path = _create_yaml_file(tmp_path, content)
        errors = validate.validate_template_file_integrity(yaml_path)
        assert len(errors) == 0


class TestValidateTemplateFileLoading:
    def test_error_when_file_does_not_exist(self):
        errors = validate.validate_template_file_loading("not a file")
        assert len(errors) == 1
        assert errors[0].error_level == validate.ErrorLevel.CRITICAL

    def test_error_when_file_is_not_a_parsable_yaml_file(self, tmp_path):
        content = '"Invalid syntax causes a parser error'
        yaml_path = _create_yaml_file(tmp_path, content)
        errors = validate.validate_template_file_loading(yaml_path)
        assert len(errors) == 1
        assert errors[0].error_level == validate.ErrorLevel.CRITICAL

    def test_no_error_when_file_is_a_parsable_yaml_file(self, tmp_path):
        content = "Key1: Value1\nKey2: Value2"
        yaml_path = _create_yaml_file(tmp_path, content)
        errors = validate.validate_template_file_loading(yaml_path)
        assert len(errors) == 0


class TestValidateTemplateDocumentRootType:
    def test_no_error_when_document_root_is_a_dictionary(self):
        errors = validate.validate_template_document_root_type({})
        assert len(errors) == 0

    def test_error_when_document_root_is_not_a_dictionary(self):
        errors = validate.validate_template_document_root_type("not a dictionary")
        assert len(errors) == 1
        assert errors[0].error_type == validate.ValidationErrorType.TYPE_ERROR
        assert errors[0].error_level == validate.ErrorLevel.CRITICAL
