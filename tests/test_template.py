import os

import pytest
import yaml

from xlsxreport.template import ReportTemplate


@pytest.fixture()
def template_document():
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
def temp_template_path(request, tmp_path):
    output_path = os.path.join(tmp_path, "template_save.yaml")

    def teardown():
        if os.path.isfile(output_path):
            os.remove(output_path)

    request.addfinalizer(teardown)
    return output_path


@pytest.fixture()
def temp_template_file(request, tmp_path, template_document):
    template_path = os.path.join(tmp_path, "template_load.yaml")

    def teardown():
        if os.path.isfile(template_path):
            os.remove(template_path)

    with open(template_path, "w", encoding="utf-8") as file:
        yaml.safe_dump(template_document, file, version=(1, 2))
    request.addfinalizer(teardown)
    return template_path


class TestReportTemplate:
    def test_from_dict_creates_template(self, template_document):
        template = ReportTemplate.from_dict(template_document)
        assert template.sections == template_document["sections"]
        assert template.formats == template_document["formats"]
        assert template.conditional_formats == template_document["conditional_formats"]  # fmt: skip
        assert template.settings == template_document["settings"]

    def test_to_dict_returns_template_document(self, template_document):
        template = ReportTemplate(
            sections=template_document["sections"],
            formats=template_document["formats"],
            conditional_formats=template_document["conditional_formats"],
            settings=template_document["settings"],
        )
        assert template.to_dict() == template_document

    def test_round_trip_from_dict_to_dict(self, template_document):
        template = ReportTemplate.from_dict(template_document)
        assert template.to_dict() == template_document

    def test_load_imports_all_sections_properly(self, temp_template_file, template_document):  # fmt: skip
        template = ReportTemplate.load(temp_template_file)
        assert template.sections == template_document["sections"]
        assert template.formats == template_document["formats"]
        assert template.conditional_formats == template_document["conditional_formats"]
        assert template.settings == template_document["settings"]

    def test_template_identical_after_load_save_reload(self, temp_template_file, temp_template_path):  # fmt: skip
        template = ReportTemplate.load(temp_template_file)
        template.save(temp_template_path)
        loaded_template = ReportTemplate.load(temp_template_file)
        assert template.sections == loaded_template.sections
        assert template.formats == loaded_template.formats
        assert template.conditional_formats == loaded_template.conditional_formats
        assert template.settings == loaded_template.settings
