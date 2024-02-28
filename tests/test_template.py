import os

import pytest
import yaml

from xlsxreport.template import ReportTemplate


@pytest.fixture()
def default_template():
    return {
        "sections": {
            "section_1": {
                "format": "str",
                "width": 70,
                "columns": ["Column 1", "Column 2"],
                "column_format": {"Column 1": "int"},
                "column_conditional": {"Column 1": "conditional"},
            },
            "section_2": {"columns": ["Column 3"]},
        },
        "formats": {
            "int": {"align": "center", "num_format": "0"},
            "str": {"align": "left", "num_format": "0"},
            "header": {"bold": True, "align": "center"},
        },
        "conditional_formats": {
            "conditional": {
                "type": "2_color_scale",
                "min_color": "#ffffbf",
                "max_color": "#f25540",
            }
        },
        "settings": {"header_height": 95, "column_width": 45},
    }


@pytest.fixture()
def default_template_path(tmp_path, default_template):
    file_path = tmp_path / "temp_template.yaml"
    with open(file_path, "w", encoding="utf-8") as file:
        yaml.safe_dump(default_template, file, version=(1, 2))
    return file_path


def create_yaml_from_string(tmp_path, content: str = ""):
    file_path = tmp_path / "temp_template.yaml"
    with open(file_path, "w", encoding="utf-8") as file:
        file.write(content)
    return file_path


class TestReportTemplate:
    def test_to_dict_returns_correct_template_document(self, default_template):
        template = ReportTemplate(
            sections=default_template["sections"],
            formats=default_template["formats"],
            conditional_formats=default_template["conditional_formats"],
            settings=default_template["settings"],
        )
        assert template.to_dict() == default_template

    def test_from_dict_to_dict_roundtrip(self, default_template):
        template = ReportTemplate.from_dict(default_template)
        assert template.to_dict() == default_template

    def test_load_imports_all_sections_properly(self, default_template_path, default_template):  # fmt: skip
        template = ReportTemplate.load(default_template_path)
        assert template.to_dict() == default_template

    def test_template_identical_after_load_save_reload(self, default_template_path, tmp_path):  # fmt: skip
        template = ReportTemplate.load(default_template_path)
        saved_template_path = tmp_path / "template_save.yaml"
        template.save(saved_template_path)
        loaded_template = ReportTemplate.load(saved_template_path)
        assert template.to_dict() == loaded_template.to_dict()

    def test_init_raises_value_error_when_invalid_parameters_are_passed(self):
        with pytest.raises(ValueError):
            _ = ReportTemplate(sections="not a dictionary")

    def test_from_dict_with_invalid_template_document_raises_value_error(self):
        with pytest.raises(ValueError):
            _ = ReportTemplate.from_dict({"sections": "not a dictionary"})

    @pytest.mark.parametrize(
        "content",
        ['"Parser error', "['not a', 'dictrionary']", "sections: not a dictionary"],
    )
    def test_loading_an_invalid_yaml_file_raises_a_value_error(self, content, tmp_path):
        yaml_path = create_yaml_from_string(tmp_path, content)
        with pytest.raises(ValueError):
            _ = ReportTemplate.load(yaml_path)
