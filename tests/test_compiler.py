import pytest
import pandas as pd
from sympy import comp
import xlsxreport.compiler as compiler
from xlsxreport.template import ReportTemplate


@pytest.fixture()
def report_template() -> ReportTemplate:
    section_template = {
        "format": "str",
        "columns": ["Column 1", "Column 2", "Column 3"],
        "column_format": {"Column 1": "float"},
        "column_conditional": {"Column 1": "cond_1"},
        "supheader": "Supheader",
        "conditional": "cond_2",
    }

    report_template = ReportTemplate(
        sections={"section 1": section_template},
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
        settings={"column_width": 10},
    )
    return report_template


@pytest.fixture()
def example_table() -> pd.DataFrame:
    table = pd.DataFrame(
        {
            "Column 1": [1, 2, 3],
            "Column 2": ["A", "B", "C"],
            "Column 4": ["A", "B", "C"],
        }
    )
    return table


@pytest.fixture()
def table_section(report_template, example_table) -> compiler.TableSection:
    table_section = compiler.TableSection(
        data=example_table[["Column 1", "Column 2"]].copy(),
        column_formats={
            "Column 1": report_template.formats["float"],
            "Column 2": report_template.formats["str"],
        },
        column_conditionals={
            "Column 1": report_template.conditional_formats["cond_1"],
            "Column 2": {},
        },
        column_widths={
            "Column 1": 10,
            "Column 2": 10,
        },
        header_formats={
            "Column 1": report_template.formats["header"],
            "Column 2": report_template.formats["header"],
        },
        supheader=report_template.sections["section 1"]["supheader"],
        supheader_format=report_template.formats["supheader"],
        section_conditional=report_template.conditional_formats["cond_2"],
    )
    return table_section


def test_eval_section_columns_selects_correct_columns():
    template_section = {"columns": ["Column 1", "Column 2", "Column 3"]}
    table = pd.DataFrame(columns=["Column 1", "Column 2", "Column 4"])
    selected_columns = compiler.eval_section_columns(template_section, table)
    assert selected_columns == ["Column 1", "Column 2"]


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


class TestEvalColumnConditionalFormats:
    @pytest.fixture(autouse=True)
    def _init_inputs(self):
        self.columns = ["Column 1", "Column 2"]
        self.section_template = {"column_conditional": {"Column 1": "cond_1"}}
        self.conditional_format_templates = {"cond_1": {"type": "2_color_scale"}}

    def test_with_no_column_conditional_defined_in_the_section_template(self):
        column_formats = compiler.eval_column_conditional_formats(
            self.columns, {}, self.conditional_format_templates
        )
        expected_formats = {"Column 1": {}, "Column 2": {}}
        assert column_formats == expected_formats

    def test_with_column_conditional_defined_in_the_section_template(self):
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
        self.section_template["column_conditional"] = {"Column 1": "cond_1", "Column 2": "cond_1"}  # fmt: skip
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


class TestEvalSectionConditionalFormats:
    @pytest.fixture(autouse=True)
    def _init_inputs(self):
        self.section_template = {"conditional": "cond_1"}
        self.conditional_format_templates = {"cond_1": {"type": "2_color_scale"}}

    def test_with_no_column_conditional_defined_in_the_section_template(self):
        column_formats = compiler.eval_section_conditional_format({}, self.conditional_format_templates)  # fmt: skip
        expected_formats = {}
        assert column_formats == expected_formats

    def test_with_column_conditional_defined_in_the_section_template(self):
        column_formats = compiler.eval_section_conditional_format(self.section_template, self.conditional_format_templates)  # fmt: skip
        expected_formats = self.conditional_format_templates["cond_1"]
        assert column_formats == expected_formats

    def test_with_no_conditional_defined_in_conditional_format_templates(self):
        column_formats = compiler.eval_section_conditional_format(self.section_template, {})  # fmt: skip
        expected_formats = {}
        assert column_formats == expected_formats


# Tests are missing for several edge cases, especially related to default formats
# - currently it is not tested when no width is specified in the ReportTemplate
def test_compile_table_section(report_template, table_section, example_table):
    compiled_sections = compiler.compile_table_sections(report_template, example_table)
    compiled_section = compiled_sections[0]
    for attr in table_section.__dataclass_fields__:
        if attr == "data":
            assert compiled_section.data.equals(table_section.data)
        else:
            # Include attribute name in a dictionary to get nicer error messages
            compiled_attr = {attr: getattr(compiled_section, attr)}
            table_section_attr = {attr: getattr(table_section, attr)}
            assert compiled_attr == table_section_attr
