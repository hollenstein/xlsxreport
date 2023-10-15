import pytest
import os
from xlsxreport.config import ReportConfig


@pytest.fixture()
def temp_config_path(request, tmp_path):
    output_path = os.path.join(tmp_path, "config_save.yaml")

    def teardown():
        if os.path.isfile(output_path):
            os.remove(output_path)

    request.addfinalizer(teardown)
    return output_path


@pytest.fixture()
def temp_config_file(request, tmp_path):
    config_path = os.path.join(tmp_path, "config_load.yaml")

    def teardown():
        if os.path.isfile(config_path):
            os.remove(config_path)

    config_text = """
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
    config_text = "\n".join([line[8:] for line in config_text.splitlines()]).strip()
    with open(config_path, "w", encoding="utf-8") as file:
        file.write(config_text)
    request.addfinalizer(teardown)
    return config_path


class TestReportConfig:
    def test_config_identical_after_load_save_reload(
        self, temp_config_file, temp_config_path
    ):
        config = ReportConfig.load(temp_config_file)
        config.save(temp_config_path)
        loaded_config = ReportConfig.load(temp_config_file)
        assert config.groups == loaded_config.groups
        assert config.formats == loaded_config.formats
        assert config.conditional_formats == loaded_config.conditional_formats
        assert config.settings == loaded_config.settings
