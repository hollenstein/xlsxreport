import numpy as np
import pytest
import pandas as pd

import xlsxreport.compiler as compiler
from xlsxreport.template import TableTemplate
from xlsxreport.template import TemplateSection


@pytest.fixture()
def table_template() -> TableTemplate:
    standard_section_template = {
        "format": "str",
        "columns": ["Column 1", "Column 2", "Column 3"],
        "column_format": {"Column 1": "float"},
        "column_conditional_format": {"Column 1": "cond_1"},
        "supheader": "Supheader",
        "conditional_format": "cond_2",
    }

    tag_sample_section_template = {
        "tag": "Tag",
        "format": "float",
        "supheader": "Supheader",
        "conditional_format": "cond_1",
        "remove_tag": True,
        "log2": True,
    }

    label_tag_sample_section_template = {
        "tag": "Tag",
        "labels": ["Sample 1"],
        "format": "float",
        "supheader": "Supheader",
        "conditional_format": "cond_1",
        "remove_tag": True,
        "log2": True,
    }

    table_template = TableTemplate(
        sections={
            "Standard section 1": standard_section_template,
            "Tag section 1": tag_sample_section_template,
            "Label tag section 1": label_tag_sample_section_template,
        },
        formats={
            "header": {"bold": True, "align": "center"},
            "supheader": {"bold": True, "align": "center", "text_wrap": True},
            "str": {"align": "left", "num_format": "@"},
            "float": {"align": "center", "num_format": "0.00"},
        },
        conditional_formats={
            "cond_1": {"type": "2_color_scale"},
            "cond_2": {"type": "3_color_scale"},
        },
        settings={
            "column_width": 10,
            "log2_tag": "[log2]",
            "remove_duplicate_columns": True,
        },
    )
    return table_template


@pytest.fixture()
def example_table() -> pd.DataFrame:
    example_table = pd.DataFrame(
        {
            "Column 1": [1, 2, 3],
            "Column 2": ["A", "B", "C"],
            "Column 4": ["A", "B", "C"],
            "Tag Sample 1": [1, 2, 3],
            "Tag Sample 2": [1, 2, 3],
        }
    )
    return example_table


@pytest.fixture()
def compiled_standard_section(table_template, example_table) -> compiler.CompiledSection:  # fmt: skip
    compiled_section = compiler.CompiledSection(
        data=example_table[["Column 1", "Column 2"]].copy(),
        column_formats={
            "Column 1": table_template.formats["float"],
            "Column 2": table_template.formats["str"],
        },
        column_conditional_formats={
            "Column 1": table_template.conditional_formats["cond_1"],
            "Column 2": {},
        },
        column_widths={
            "Column 1": 10,
            "Column 2": 10,
        },
        headers={"Column 1": "Column 1", "Column 2": "Column 2"},
        header_formats={
            "Column 1": table_template.formats["header"],
            "Column 2": table_template.formats["header"],
        },
        supheader=table_template.sections["Standard section 1"]["supheader"],
        supheader_format=table_template.formats["supheader"],
        section_conditional_format=table_template.conditional_formats["cond_2"],
    )
    return compiled_section


@pytest.fixture()
def compiled_label_tag_sample_section(table_template, example_table) -> compiler.CompiledSection:  # fmt: skip
    compiled_section = compiler.CompiledSection(
        data=np.log2(example_table[["Tag Sample 1"]].copy()),
        column_formats={
            "Tag Sample 1": table_template.formats["float"],
        },
        column_conditional_formats={
            "Tag Sample 1": {},
        },
        column_widths={
            "Tag Sample 1": 10,
        },
        header_formats={
            "Tag Sample 1": table_template.formats["header"],
        },
        headers={"Tag Sample 1": "Sample 1"},
        supheader="Supheader [log2]",
        supheader_format=table_template.formats["supheader"],
        section_conditional_format=table_template.conditional_formats["cond_1"],
    )
    return compiled_section


@pytest.fixture()
def compiled_tag_section(table_template, example_table) -> compiler.CompiledSection:
    compiled_section = compiler.CompiledSection(
        data=np.log2(example_table[["Tag Sample 1", "Tag Sample 2"]].copy()),
        column_formats={
            "Tag Sample 1": table_template.formats["float"],
            "Tag Sample 2": table_template.formats["float"],
        },
        column_conditional_formats={
            "Tag Sample 1": {},
            "Tag Sample 2": {},
        },
        column_widths={
            "Tag Sample 1": 10,
            "Tag Sample 2": 10,
        },
        header_formats={
            "Tag Sample 1": table_template.formats["header"],
            "Tag Sample 2": table_template.formats["header"],
        },
        headers={"Tag Sample 1": "Sample 1", "Tag Sample 2": "Sample 2"},
        supheader="Supheader [log2]",
        supheader_format=table_template.formats["supheader"],
        section_conditional_format=table_template.conditional_formats["cond_1"],
    )
    return compiled_section


class TestEvalData:
    def test_data_frame_contains_only_selected_columns(self):
        table = pd.DataFrame({"Column 1": [1, 2, None], "Column 2": ["A", "B", "C"]})
        evaluated_data = compiler.eval_data(table, ["Column 1"])
        assert list(evaluated_data.columns) == ["Column 1"]

    def test_nan_values_in_dataframe_replaced(self):
        table = pd.DataFrame({"Column 1": [1, 2, None], "Column 2": ["A", "B", "C"]})
        evaluated_data = compiler.eval_data(table, ["Column 1"])
        assert not evaluated_data.isna().values.any()
        assert evaluated_data["Column 1"][2] == compiler.NAN_REPLACEMENT_SYMBOL


class TestEvalDataWithLog2Transformation:
    def test_log2_transformation_applied_when_specified(self):
        table = pd.DataFrame({"Column 1": [1, 2, 3], "Column 2": ["A", "B", "C"]})
        evaluated_data = compiler.eval_data_with_log2_transformation(
            table, ["Column 1"], {"log2": True}, evaluate_log_state=False
        )
        assert evaluated_data["Column 1"].tolist() == np.log2(table["Column 1"]).tolist()  # fmt: skip

    def test_log2_transformation_replaces_values_smaller_or_equal_to_zero_with_nan(self):  # fmt: skip
        table = pd.DataFrame({"Column 1": [-1, 0, 1], "Column 2": ["A", "B", "C"]})
        evaluated_data = compiler.eval_data_with_log2_transformation(
            table, ["Column 1"], {"log2": True}, evaluate_log_state=False
        )
        assert evaluated_data["Column 1"].tolist() == ["", "", 0]

    def test_log_space_evaluation_prevents_log2_transformation_of_small_values(self):  # fmt: skip
        table = pd.DataFrame({"Column 1": [1, 2, 3], "Column 2": ["A", "B", "C"]})
        evaluated_data = compiler.eval_data_with_log2_transformation(
            table, ["Column 1"], {"log2": True}, evaluate_log_state=True
        )
        assert evaluated_data["Column 1"].tolist() == [1, 2, 3]

    def test_log_space_evaluation_does_not_prevent_log2_transformation_of_large_values(self):  # fmt: skip
        table = pd.DataFrame({"Column 1": [1, 2, 65], "Column 2": ["A", "B", "C"]})
        evaluated_data = compiler.eval_data_with_log2_transformation(
            table, ["Column 1"], {"log2": True}, evaluate_log_state=True
        )
        assert evaluated_data["Column 1"].tolist() == np.log2(table["Column 1"]).tolist()  # fmt:skip


def test_eval_standard_section_columns_selects_correct_columns():
    section_template = {"columns": ["Column 1", "Column 2", "Column 3"]}
    columns = ["Column 1", "Column 2", "Column 4"]
    selected_columns = compiler.eval_standard_section_columns(columns, section_template)
    assert selected_columns == ["Column 1", "Column 2"]


@pytest.mark.parametrize(
    "tag, expected_selection",
    [
        ("Tag", ["Tag S1", "Tag S2", "S3 Tag", "Tag"]),
        ("^Tag", ["Tag S1", "Tag S2", "Tag"]),
        ("^Tag.", ["Tag S1", "Tag S2"]),
        (".Tag$", ["S3 Tag"]),
        ("^Tag.|.Tag$", ["Tag S1", "Tag S2", "S3 Tag"]),
    ],
)
def test_eval_tag_section_columns(tag, expected_selection):
    columns = ["Tag S1", "Tag S2", "S3 Tag", "Tag", "Col1", "Col2"]
    section_template = {"tag": tag}
    selected_columns = compiler.eval_tag_section_columns(columns, section_template)
    assert selected_columns == expected_selection


class TestEvalLabelTagSectionColumns:
    @pytest.mark.parametrize(
        "tag, labels, expected_selection",
        [
            ("Tag", ["S1", "S2", "S3"], ["Tag S1", "Tag S2", "S3 Tag"]),
            ("^Tag", ["S1", "S2", "S3"], ["Tag S1", "Tag S2"]),
            ("Tag", ["S2"], ["Tag S2"]),
            (".Tag$", ["S1", "S2", "S3"], ["S3 Tag"]),
            (".Tag$", ["S1", "S2"], []),
        ],
    )
    def test_correct_columns_are_selected(self, tag, labels, expected_selection):
        columns = ["Tag S1", "Tag S2", "S3 Tag", "Tag", "Col1", "Col2"]
        section_template = {"tag": tag, "labels": labels}
        selected_columns = compiler.eval_label_tag_section_columns(columns, section_template)  # fmt: skip
        assert selected_columns == expected_selection

    def test_that_labels_determine_column_order(self):
        columns = ["Tag S1", "Tag S2", "Tag S3", "Tag", "Col1", "Col2"]
        labels = ["S3", "S1", "S2"]
        section_template = {"tag": "Tag", "labels": labels}
        selected_columns = compiler.eval_label_tag_section_columns(columns, section_template)  # fmt: skip
        assert selected_columns == ["Tag S3", "Tag S1", "Tag S2"]


def test_eval_comparison_groups_extracts_correct_values():
    section_template = {
        "comparison_group": True,
        "tag": " vs ",
        "columns": ["P", "A"],
    }
    columns = ["P", "A", "P ex1 vs ex2", "A ex1 vs ex2", "P ex1 vs EX3", "A ex1 vs EX3"]
    comparison_groups = compiler.eval_comparison_groups(columns, section_template)
    assert comparison_groups == ["ex1 vs ex2", "ex1 vs EX3"]


class TestEvalComparisonGroupColumns:
    def test_correct_columns_are_selected(self):
        section_template = {
            "comparison_group": True,
            "tag": " vs ",
            "columns": ["P", "A"],
        }
        columns = ["P", "A", "P ex1 vs ex2", "A ex1 vs ex2", "P ex1 vs EX3", "A ex1 vs EX3"]  # fmt: skip
        selected_columns = compiler.eval_comparison_group_columns(columns, section_template, "ex1 vs EX3")  # fmt: skip
        expected_columns = ["P ex1 vs EX3", "A ex1 vs EX3"]
        assert selected_columns == expected_columns

    def test_column_order_of_section_template_is_used(self):
        section_template = {
            "comparison_group": True,
            "tag": " vs ",
            "columns": ["P", "A"],
        }
        columns = ["A ex1 vs ex2", "P ex1 vs ex2"]  # fmt: skip
        selected_columns = compiler.eval_comparison_group_columns(columns, section_template, "ex1 vs ex2")  # fmt: skip
        expected_columns = ["P ex1 vs ex2", "A ex1 vs ex2"]
        assert selected_columns == expected_columns


class TestEvalComparisonGroupHeaders:
    @pytest.fixture(autouse=True)
    def _init_inputs(self):
        self.columns = ["P ex1 vs ex2", "A ex1 vs ex2", "C ex1 vs ex2"]
        self.comparison_group = "ex1 vs ex2"
        self.section_template = {"tag": " vs "}

    def test_by_default_columns_are_headers(self):
        headers = compiler.eval_comparison_group_headers(
            self.columns, self.section_template, self.comparison_group
        )
        assert headers == {c: c for c in self.columns}

    def test_that_remove_tag_removes_the_comparison_group(self):
        self.section_template["remove_tag"] = True
        headers = compiler.eval_comparison_group_headers(
            self.columns, self.section_template, self.comparison_group
        )
        expected_headers = {"P ex1 vs ex2": "P", "A ex1 vs ex2": "A", "C ex1 vs ex2": "C"}  # fmt: skip
        assert headers == expected_headers

    def test_that_replace_comparison_tag_modifies_the_header(self):
        self.section_template["replace_comparison_tag"] = " / "
        headers = compiler.eval_comparison_group_headers(
            self.columns, self.section_template, self.comparison_group
        )
        expected_headers = {
            "P ex1 vs ex2": "P ex1 / ex2",
            "A ex1 vs ex2": "A ex1 / ex2",
            "C ex1 vs ex2": "C ex1 / ex2",
        }
        assert headers == expected_headers


@pytest.mark.parametrize(
    "replace_comparison_tag, expected_supheader",
    [(None, "A vs B"), (True, "A / B")],
)
def test_eval_comparison_group_supheader(replace_comparison_tag, expected_supheader):
    section_template = {"tag": " vs "}
    if replace_comparison_tag is not None:
        section_template["replace_comparison_tag"] = " / "
    supheader = compiler.eval_comparison_group_supheader(section_template, "A vs B")
    assert supheader == expected_supheader


def test_eval_comparison_group_conditional_format_names():
    columns = ["P ex1 vs ex2", "A ex1 vs ex2", "C ex1 vs ex2"]
    section_template = {"column_conditional_format": {"P": "cond_1", "A": "cond_2"}}
    col_conditionals = compiler.eval_comparison_group_conditional_format_names(columns, section_template)  # fmt: skip
    assert col_conditionals == {"P ex1 vs ex2": "cond_1", "A ex1 vs ex2": "cond_2"}


class TestEvalTagSampleHeaders:
    def test_with_remove_tag(self):
        section_template = {"tag": "Tag", "remove_tag": True}
        columns = ["Tag Sample 1", "Tag Sample 2"]
        expected = {"Tag Sample 1": "Sample 1", "Tag Sample 2": "Sample 2"}
        headers = compiler.eval_tag_sample_headers(columns, section_template)
        assert headers == expected

    def test_without_remove_tag(self):
        section_template = {"tag": "Tag", "remove_tag": False}
        columns = ["Tag Sample 1", "Tag Sample 2"]
        expected = {"Tag Sample 1": "Tag Sample 1", "Tag Sample 2": "Tag Sample 2"}
        headers = compiler.eval_tag_sample_headers(columns, section_template)
        assert headers == expected

    def test_without_remove_tag_and_log2_tag(self):
        section_template = {"tag": "Tag", "remove_tag": False, "log2": True}
        columns = ["Tag Sample 1", "Tag Sample 2"]
        expected = {
            "Tag Sample 1": "Tag Sample 1 [log2]",
            "Tag Sample 2": "Tag Sample 2 [log2]",
        }
        headers = compiler.eval_tag_sample_headers(columns, section_template, log2_tag="[log2]")  # fmt: skip
        assert headers == expected

    def test_with_remove_tag_and_log2_tag(self):
        section_template = {"tag": "Tag", "remove_tag": True, "log2": True}
        columns = ["Tag Sample 1", "Tag Sample 2"]
        expected = {
            "Tag Sample 1": "Sample 1",
            "Tag Sample 2": "Sample 2",
        }
        headers = compiler.eval_tag_sample_headers(columns, section_template, log2_tag="[log2]")  # fmt: skip
        assert headers == expected

    def test_empty_log2_tag_does_not_modify_header(self):
        section_template = {"tag": "Tag", "remove_tag": False, "log2": True}
        columns = ["Tag Sample 1", "Tag Sample 2"]
        expected = {
            "Tag Sample 1": "Tag Sample 1",
            "Tag Sample 2": "Tag Sample 2",
        }
        headers = compiler.eval_tag_sample_headers(columns, section_template, log2_tag="")  # fmt: skip
        assert headers == expected


class TestEvalTagSampleSupheader:
    def test_with_log2_tag(self):
        section_template = {"supheader": "Supheader", "log2": True}
        supheader = compiler.eval_tag_sample_supheader(section_template, log2_tag="[log2]")  # fmt: skip
        assert supheader == "Supheader [log2]"

    def test_without_log2_tag(self):
        section_template = {"supheader": "Supheader", "log2": False}
        supheader = compiler.eval_tag_sample_supheader(section_template, log2_tag="[log2]")  # fmt: skip
        assert supheader == "Supheader"

    def test_empty_log2_tag_does_not_modify_supheader(self):
        section_template = {"supheader": "Supheader", "log2": True}
        supheader = compiler.eval_tag_sample_supheader(section_template, log2_tag="")  # fmt: skip
        assert supheader == "Supheader"


class TestEvalColumnFormats:
    @pytest.fixture(autouse=True)
    def _init_inputs(self):
        self.columns = ["Column 1", "Column 2"]
        self.section_template = {"format": "str", "column_format": {"Column 1": "float"}}  # fmt: skip
        self.format_templates = {"str": {"align": "left"}, "float": {"align": "center"}}

    def test_with_only_section_format_specified(self):
        column_formats = compiler.eval_column_formats(
            self.columns, {"format": "str"}, self.format_templates
        )
        expected_formats = {
            "Column 1": self.format_templates["str"],
            "Column 2": self.format_templates["str"],
        }
        assert column_formats == expected_formats

    def test_with_section_and_column_format_specified(self):
        column_formats = compiler.eval_column_formats(
            self.columns, self.section_template, self.format_templates
        )
        expected_formats = {
            "Column 1": self.format_templates["float"],
            "Column 2": self.format_templates["str"],
        }
        assert column_formats == expected_formats

    def test_with_no_section_or_column_format_specified(self):
        column_formats = compiler.eval_column_formats(
            self.columns, {}, self.format_templates, default_format={"test": "test"}
        )
        expected_formats = {
            "Column 1": {"test": "test"},
            "Column 2": {"test": "test"},
        }
        assert column_formats == expected_formats

    def test_that_each_returned_format_is_a_unique_instance(self):
        column_formats = compiler.eval_column_formats(
            self.columns, {"format": "str"}, self.format_templates
        )
        assert column_formats["Column 1"] is not column_formats["Column 2"]
        assert column_formats["Column 1"] == column_formats["Column 2"]

    def test_border_true_adds_right_and_left_border_only_to_first_and_last_column(self):  # fmt: skip
        columns = ["Col 1", "Col 2", "Col 3"]
        column_formats = compiler.eval_column_formats(columns, {"border": True}, {})
        assert column_formats[columns[0]] == {"left": compiler.BORDER_TYPE}
        assert column_formats[columns[1]] == {}
        assert column_formats[columns[-1]] == {"right": compiler.BORDER_TYPE}

    def test_border_true_with_one_column_adds_left_and_right_border(self):
        self.section_template["border"] = True
        column_formats = compiler.eval_column_formats(["Col 1"], {"border": True}, {})
        assert column_formats["Col 1"] == {"left": compiler.BORDER_TYPE, "right": compiler.BORDER_TYPE}  # fmt: skip


class TestEvalColumnConditionalFormats:
    @pytest.fixture(autouse=True)
    def _init_inputs(self):
        self.columns = ["Column 1", "Column 2"]
        self.section_template = {"column_conditional_format": {"Column 1": "cond_1"}}
        self.conditional_format_templates = {"cond_1": {"type": "2_color_scale"}}

    def test_with_no_column_conditional_format_defined_in_the_section_template(self):
        column_formats = compiler.eval_column_conditional_formats(
            self.columns, {}, self.conditional_format_templates
        )
        expected_formats = {"Column 1": {}, "Column 2": {}}
        assert column_formats == expected_formats

    def test_with_column_conditional_format_defined_in_the_section_template(self):
        column_formats = compiler.eval_column_conditional_formats(
            self.columns, self.section_template, self.conditional_format_templates
        )
        expected_formats = {
            "Column 1": self.conditional_format_templates["cond_1"],
            "Column 2": {},
        }
        assert column_formats == expected_formats

    def test_with_no_conditional_defined_in_conditional_format_templates(self):
        column_formats = compiler.eval_column_conditional_formats(self.columns, self.section_template, {})  # fmt: skip
        expected_formats = {"Column 1": {}, "Column 2": {}}
        assert column_formats == expected_formats

    def test_that_each_returned_format_is_a_unique_instance(self):
        self.section_template["column_conditional_format"] = {"Column 1": "cond_1", "Column 2": "cond_1"}  # fmt: skip
        column_formats = compiler.eval_column_conditional_formats(
            self.columns, self.section_template, self.conditional_format_templates
        )
        assert column_formats["Column 1"] is not column_formats["Column 2"]
        assert column_formats["Column 1"] == column_formats["Column 2"]


class TestEvalColumnWidths:
    @pytest.fixture(autouse=True)
    def _init_inputs(self):
        self.columns = ["Column 1", "Column 2"]
        self.section_template = {"width": 70}

    def test_with_section_width_set(self):
        column_widths = compiler.eval_column_widths(self.columns, self.section_template)
        expected_widths = {"Column 1": 70, "Column 2": 70}
        assert column_widths == expected_widths

    def test_with_no_section_width_set(self):
        column_widths = compiler.eval_column_widths(self.columns, {}, default_width=0)
        expected_widths = {"Column 1": 0, "Column 2": 0}
        assert column_widths == expected_widths


class TestEvalHeaderFormats:
    @pytest.fixture(autouse=True)
    def _init_inputs(self):
        self.columns = ["Column 1", "Column 2"]
        self.section_template = {"header_format": {"bold": True}}
        self.format_templates = {"header": {"align": "center"}}

    def test_with_header_format_specified_in_section_and_format_template(self):
        header_formats = compiler.eval_header_formats(
            self.columns, self.section_template, self.format_templates
        )
        assert header_formats == {col: {"align": "center", "bold": True} for col in self.columns}  # fmt: skip

    def test_with_header_format_specified_in_section_but_not_in_format_template(self):
        header_formats = compiler.eval_header_formats(self.columns, self.section_template, {})  # fmt: skip
        assert header_formats == {col: {"bold": True} for col in self.columns}

    def test_with_header_format_not_specified_in_section_but_in_format_template(self):
        header_formats = compiler.eval_header_formats(self.columns, {}, self.format_templates)  # fmt: skip
        assert header_formats == {col: {"align": "center"} for col in self.columns}

    def test_with_header_format_not_specified_in_section_and_not_in_format_template(self):  # fmt: skip
        header_formats = compiler.eval_header_formats(self.columns, {}, {})
        assert header_formats == {col: {} for col in self.columns}

    def test_that_header_format_in_section_overwrites_the_header_format_template(self):
        header_formats = compiler.eval_header_formats(
            self.columns, {"header_format": {"a": 1}}, {"header": {"a": 2}}
        )
        assert header_formats == {col: {"a": 1} for col in self.columns}

    def test_that_each_returned_header_format_is_a_unique_instance(self):
        column_formats = compiler.eval_header_formats(
            self.columns, self.section_template, self.format_templates
        )
        assert column_formats["Column 1"] is not column_formats["Column 2"]
        assert column_formats["Column 1"] == column_formats["Column 2"]

    def test_that_original_formats_are_not_overwritten(self):
        original_template_format = self.format_templates["header"].copy()
        original_section_format = self.section_template["header_format"].copy()
        _ = compiler.eval_header_formats(
            self.columns, self.section_template, self.format_templates
        )
        assert original_template_format == self.format_templates["header"]
        assert original_section_format == self.section_template["header_format"]

    def test_border_true_adds_right_and_left_border_only_to_first_and_last_column(self):  # fmt: skip
        columns = ["Col 1", "Col 2", "Col 3"]
        header_formats = compiler.eval_header_formats(columns, {"border": True}, {})
        assert header_formats[columns[0]] == {"left": compiler.BORDER_TYPE}
        assert header_formats[columns[1]] == {}
        assert header_formats[columns[-1]] == {"right": compiler.BORDER_TYPE}

    def test_border_true_with_one_column_adds_left_and_right_border(self):
        self.section_template["border"] = True
        header_formats = compiler.eval_header_formats(
            ["Col 1"], {"format": "str", "border": True}, {"str": {}}
        )
        assert header_formats["Col 1"] == {"left": compiler.BORDER_TYPE, "right": compiler.BORDER_TYPE}  # fmt: skip


class TestEvalSupHeaderFormat:
    @pytest.fixture(autouse=True)
    def _init_inputs(self):
        self.section_template = {"supheader_format": {"bold": True}}
        self.format_templates = {"supheader": {"align": "center"}}

    def test_with_format_specified_in_section_and_format_template(self):
        sup_format = compiler.eval_supheader_format(self.section_template, self.format_templates)  # fmt: skip
        assert sup_format == {"align": "center", "bold": True}

    def test_with_format_specified_in_section_but_not_in_format_template(self):
        sup_format = compiler.eval_supheader_format(self.section_template, {})
        assert sup_format == {"bold": True}

    def test_with_format_not_specified_in_section_but_in_format_template(self):
        sup_format = compiler.eval_supheader_format({}, self.format_templates)
        assert sup_format == {"align": "center"}

    def test_with_format_not_specified_in_section_and_not_in_format_template(self):
        sup_format = compiler.eval_supheader_format({}, {})
        assert sup_format == {}

    def test_that_format_in_section_overwrites_the_header_format_template(self):
        sup_format = compiler.eval_supheader_format({"supheader_format": {"a": 1}}, {"supheader": {"a": 2}})  # fmt: skip
        assert sup_format == {"a": 1}

    def test_that_original_formats_are_not_overwritten(self):
        original_template_format = self.format_templates["supheader"].copy()
        original_section_format = self.section_template["supheader_format"].copy()
        _ = compiler.eval_supheader_format(self.section_template, self.format_templates)
        assert original_template_format == self.format_templates["supheader"]
        assert original_section_format == self.section_template["supheader_format"]

    def test_border_true(self):
        sup_format = compiler.eval_supheader_format({"border": True}, {})
        assert sup_format == {"left": compiler.BORDER_TYPE, "right": compiler.BORDER_TYPE}  # fmt: skip


class TestEvalSectionConditionalFormats:
    @pytest.fixture(autouse=True)
    def _init_inputs(self):
        self.section_template = {"conditional_format": "cond_1"}
        self.conditional_format_templates = {"cond_1": {"type": "2_color_scale"}}

    def test_with_no_column_conditional_format_defined_in_the_section_template(self):
        column_formats = compiler.eval_section_conditional_format({}, self.conditional_format_templates)  # fmt: skip
        expected_formats = {}
        assert column_formats == expected_formats

    def test_with_column_conditional_format_defined_in_the_section_template(self):
        column_formats = compiler.eval_section_conditional_format(self.section_template, self.conditional_format_templates)  # fmt: skip
        expected_formats = self.conditional_format_templates["cond_1"]
        assert column_formats == expected_formats

    def test_with_no_conditional_defined_in_conditional_format_templates(self):
        column_formats = compiler.eval_section_conditional_format(self.section_template, {})  # fmt: skip
        expected_formats = {}
        assert column_formats == expected_formats


def test_StandardSectionCompiler(table_template, example_table, compiled_standard_section):  # fmt: skip
    section_compiler = compiler.StandardSectionCompiler(table_template)
    section_template = table_template.sections["Standard section 1"].to_dict()
    compiled_section = section_compiler.compile(section_template, example_table)[0]
    for attr in compiled_standard_section.__dataclass_fields__:
        if attr == "data":
            pd.testing.assert_frame_equal(compiled_section.data, compiled_standard_section.data, check_dtype=False)  # fmt: skip
        else:
            # Include attribute name in a dictionary to get nicer error messages
            compiled_attr = {attr: getattr(compiled_section, attr)}
            expected_section_attr = {attr: getattr(compiled_standard_section, attr)}
            assert compiled_attr == expected_section_attr


def test_TagSectionCompiler(table_template, example_table, compiled_tag_section):
    section_compiler = compiler.TagSectionCompiler(table_template)
    section_template = table_template.sections["Tag section 1"].to_dict()
    compiled_section = section_compiler.compile(section_template, example_table)[0]
    for attr in compiled_tag_section.__dataclass_fields__:
        if attr == "data":
            pd.testing.assert_frame_equal(compiled_section.data, compiled_tag_section.data, check_dtype=False)  # fmt: skip
        else:
            # Include attribute name in a dictionary to get nicer error messages
            compiled_attr = {attr: getattr(compiled_section, attr)}
            expected_section_attr = {attr: getattr(compiled_tag_section, attr)}
            assert compiled_attr == expected_section_attr


def test_LabelTagSectionCompiler(table_template, example_table, compiled_label_tag_sample_section):  # fmt: skip
    section_compiler = compiler.LabelTagSectionCompiler(table_template)
    section_template = table_template.sections["Label tag section 1"].to_dict()
    compiled_section = section_compiler.compile(section_template, example_table)[0]
    for attr in compiled_label_tag_sample_section.__dataclass_fields__:
        if attr == "data":
            pd.testing.assert_frame_equal(compiled_section.data, compiled_label_tag_sample_section.data, check_dtype=False)  # fmt: skip
        else:
            # Include attribute name in a dictionary to get nicer error messages
            compiled_attr = {attr: getattr(compiled_section, attr)}
            expected_section_attr = {
                attr: getattr(compiled_label_tag_sample_section, attr)
            }
            assert compiled_attr == expected_section_attr


class TestComparisonSectionCompiler:
    @pytest.fixture(autouse=True)
    def _init_inputs(self, table_template):
        self.table = pd.DataFrame(
            {
                "A column": [],
                "P ex1 vs ex2": [],
                "A ex1 vs ex2": [],
                "P ex1 vs EX3": [],
                "A ex1 vs EX3": [],
            }
        )
        self.section_template = {
            "comparison_group": True,
            "tag": " vs ",
            "columns": ["P", "A"],
            "column_conditional_format": {
                "P": "cond_1",
                "A": "cond_2",
            },
            "remove_tag": True,
            "replace_comparison_tag": " / ",
        }
        self.section_compiler = compiler.ComparisonSectionCompiler(table_template)

    def test_that_two_sections_with_correct_columns_are_generated(self):
        compiled_sections = self.section_compiler.compile(self.section_template, self.table)  # fmt: skip
        assert len(compiled_sections) == 2

        section_1, section_2 = compiled_sections
        assert section_1.data.columns.tolist() == ["P ex1 vs ex2", "A ex1 vs ex2"]
        assert section_2.data.columns.tolist() == ["P ex1 vs EX3", "A ex1 vs EX3"]

    def test_correct_application_of_conditional_formats(self):
        compiled_sections = self.section_compiler.compile(self.section_template, self.table)  # fmt: skip
        expected_column_conditional_formats = {
            "P ex1 vs ex2": {"type": "2_color_scale"},
            "A ex1 vs ex2": {"type": "3_color_scale"},
        }
        observed_column_conditional_formats = compiled_sections[0].column_conditional_formats  # fmt: skip
        assert observed_column_conditional_formats == expected_column_conditional_formats  # fmt: skip

    def test_compiled_sections_have_correct_headers(self):
        compiled_sections = self.section_compiler.compile(self.section_template, self.table)  # fmt: skip
        expected_headers = {"P ex1 vs ex2": "P", "A ex1 vs ex2": "A"}
        assert compiled_sections[0].headers == expected_headers

    def test_compiled_sections_have_correct_supheader(self):
        compiled_sections = self.section_compiler.compile(self.section_template, self.table)  # fmt: skip
        assert compiled_sections[0].supheader == "ex1 / ex2"
        assert compiled_sections[1].supheader == "ex1 / EX3"


class TestPrepareCompiledSections:
    def test_only_non_empty_sections_are_returned(self, table_template):
        table = pd.DataFrame({"Tag Sample 1": [1], "Tag Sample 2": [1]})
        compiled_sections = compiler.prepare_compiled_sections(table_template, table)
        assert all([not s.data.empty for s in compiled_sections])

    def test_duplicate_columns_are_removed(self, table_template, example_table):
        table_template.sections["Another section"] = TemplateSection({"columns": ["Column 1"]})  # fmt: skip
        compiled_sections = compiler.prepare_compiled_sections(table_template, example_table)  # fmt: skip
        observed_columns = []
        for section in compiled_sections:
            observed_columns.extend(section.data.columns)
        assert len(set(observed_columns)) == len(observed_columns)

    def test_duplicate_columns_are_kept_when_setting_is_false(self, table_template, example_table):  # fmt: skip
        table_template.sections["Another section"] = TemplateSection({"columns": ["Column 1"]})  # fmt: skip
        table_template.settings["remove_duplicate_columns"] = False
        compiled_sections = compiler.prepare_compiled_sections(table_template, example_table)  # fmt: skip
        assert compiled_sections[-1].data.columns.tolist() == ["Column 1"]

    def test_empty_sections_caused_by_removal_of_duplicate_columns_are_not_returned(
        self, table_template, example_table
    ):
        table_template.sections["Another section"] = TemplateSection({"columns": ["Column 1"]})  # fmt: skip
        compiled_sections = compiler.prepare_compiled_sections(table_template, example_table)  # fmt: skip
        assert all([not s.data.empty for s in compiled_sections])

    def test_addition_of_section_with_remaining_columns(self, table_template):
        table = pd.DataFrame({"Column 1": [1], "Another column": [1]})
        table_template.settings["append_remaining_columns"] = True
        compiled_sections = compiler.prepare_compiled_sections(table_template, table)
        assert len(compiled_sections) == 2


class TestCompileSection:
    def test_correctly_compiled_sections(
        self,
        table_template,
        example_table,
        compiled_standard_section,
        compiled_tag_section,
    ):
        compiled_sections = compiler.compile_sections(table_template, example_table)
        expected_sections = [compiled_standard_section, compiled_tag_section]
        for compiled_section, expected_section in zip(expected_sections, compiled_sections):  # fmt: skip
            for attr in expected_section.__dataclass_fields__:
                if attr == "data":
                    pd.testing.assert_frame_equal(compiled_section.data, expected_section.data, check_dtype=False)  # fmt: skip
                else:
                    # Include attribute name in a dictionary to get nicer error messages
                    compiled_attr = {attr: getattr(compiled_section, attr)}
                    expected_section_attr = {attr: getattr(expected_section, attr)}
                    assert compiled_attr == expected_section_attr

    def test_invalid_sections_are_not_compiled(self, table_template, example_table):
        section = TemplateSection({"columns": ["Column 1", "Column 2"]})
        section.category = compiler.SectionCategory.UNKNOWN
        table_template.sections = {"invalid": section}
        compiled_sections = compiler.compile_sections(table_template, example_table)
        assert len(compiled_sections) == 0


class TestCompileRemaininColumnSection:
    def test_correct_columns_selected_for_section(
        self, table_template, example_table, compiled_standard_section
    ):
        compiled_section = compiler.compile_remaining_column_section(
            table_template, [compiled_standard_section], example_table
        )
        expected_columns = [
            c for c in example_table if c not in compiled_standard_section.data
        ]
        assert not compiled_standard_section.data.columns.empty
        assert compiled_section.data.columns.tolist() == expected_columns

    def test_empty_section_returned_when_no_remaining_columns(
        self, table_template, example_table, compiled_standard_section
    ):
        example_table = example_table[compiled_standard_section.data.columns]
        compiled_section = compiler.compile_remaining_column_section(
            table_template, [compiled_standard_section], example_table
        )
        assert compiled_section.data.empty


class TestPruneCompiledSections:
    def test_duplicate_columns_removed_from_latter_sections(self):
        sections = [
            compiler.CompiledSection(data=pd.DataFrame(columns=["C1", "C2"])),
            compiler.CompiledSection(data=pd.DataFrame(columns=["C1", "C3"])),
            compiler.CompiledSection(data=pd.DataFrame(columns=["C1", "C4"])),
        ]
        compiler.prune_compiled_sections(sections)
        assert sections[0].data.columns.tolist() == ["C1", "C2"]
        assert sections[1].data.columns.tolist() == ["C3"]
        assert sections[2].data.columns.tolist() == ["C4"]

    def test_removal_of_all_columns_from_a_section_returns_an_empty_section(self):
        sections = [
            compiler.CompiledSection(data=pd.DataFrame(columns=["C1", "C2"])),
            compiler.CompiledSection(data=pd.DataFrame(columns=["C1", "C2"])),
        ]
        compiler.prune_compiled_sections(sections)
        assert sections[0].data.columns.tolist() == ["C1", "C2"]
        assert sections[1].data.columns.tolist() == []

    def test_unique_columns_not_removed(self):
        sections = [
            compiler.CompiledSection(data=pd.DataFrame(columns=["C1", "C2"])),
            compiler.CompiledSection(data=pd.DataFrame(columns=["C3", "C4"])),
        ]
        compiler.prune_compiled_sections(sections)
        for section in sections:
            assert len(section.data.columns) == 2
            assert len(sections[1].column_formats) == 2
            assert len(sections[1].column_conditional_formats) == 2
            assert len(sections[1].column_widths) == 2
            assert len(sections[1].headers) == 2
            assert len(sections[1].header_formats) == 2

    def test_column_removed_from_all_section_parameters(self):
        sections = [
            compiler.CompiledSection(data=pd.DataFrame(columns=["C1", "C2"])),
            compiler.CompiledSection(data=pd.DataFrame(columns=["C1", "C3"])),
        ]
        compiler.prune_compiled_sections(sections)
        assert sections[1].data.columns.tolist() == ["C3"]
        assert list(sections[1].column_formats) == ["C3"]
        assert list(sections[1].column_conditional_formats) == ["C3"]
        assert list(sections[1].column_widths) == ["C3"]
        assert list(sections[1].headers) == ["C3"]
        assert list(sections[1].header_formats) == ["C3"]


def test_remove_empty_compiled_sections():
    sections = [
        compiler.CompiledSection(data=pd.DataFrame({"C1": [1]})),
        compiler.CompiledSection(data=pd.DataFrame({})),
        compiler.CompiledSection(data=pd.DataFrame({"C2": [1]})),
    ]
    filtered_sections = compiler.remove_empty_compiled_sections(sections)
    assert all([not s.data.empty for s in filtered_sections])
