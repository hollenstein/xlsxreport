import pytest
import os
from xlsxreport.template import ReportTemplate


@pytest.fixture()
def temp_template_path(request, tmp_path):
    output_path = os.path.join(tmp_path, "template_save.yaml")

    def teardown():
        if os.path.isfile(output_path):
            os.remove(output_path)

    request.addfinalizer(teardown)
    return output_path


@pytest.fixture()
def temp_template_file(request, tmp_path):
    template_path = os.path.join(tmp_path, "template_load.yaml")

    def teardown():
        if os.path.isfile(template_path):
            os.remove(template_path)

    template_text = """
        %YAML 1.2
        ---
        groups:
            group_1: {
                format: "str",
                width: 70,
                columns: ["Column 1", "Column 2",],
                column_format: {"Column 1": "int",},
                column_conditional: {"Column 1": "conditional",},
            }
            group_2: {
                columns: ["Column 3"],
            }

        formats:
            int: {"align": "center", "num_format": "0"}
            str: {"align": "left", "num_format": "0"}
            header: {"bold": True, "align": "center",}

        conditional_formats:
            conditional: {
                "type": "2_color_scale", "min_color": "#ffffbf", "max_color": "#f25540"
            }

        args:
            header_height: 95
            column_width: 45
    """
    template_text = "\n".join([line[8:] for line in template_text.splitlines()]).strip()
    with open(template_path, "w", encoding="utf-8") as file:
        file.write(template_text)
    request.addfinalizer(teardown)
    return template_path


class TestReportTemplate:
    def test_config_identical_after_load_save_reload(
        self, temp_template_file, temp_template_path
    ):
        template = ReportTemplate.load(temp_template_file)
        template.save(temp_template_path)
        loaded_template = ReportTemplate.load(temp_template_file)
        assert template.groups == loaded_template.groups
        assert template.formats == loaded_template.formats
        assert template.conditional_formats == loaded_template.conditional_formats
        assert template.settings == loaded_template.settings
