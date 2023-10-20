import pytest
import pandas as pd
from sympy import comp
import xlsxreport.compiler as compiler
from xlsxreport.template import ReportTemplate


@pytest.fixture()
def report_template() -> ReportTemplate:
    standard_section_template = {
        "format": "str",
        "columns": ["Column 1", "Column 2", "Column 3"],
        "column_format": {"Column 1": "float"},
        "column_conditional": {"Column 1": "cond_1"},
        "supheader": "Supheader",
        "conditional": "cond_2",
    }

    tag_sample_section_template = {
        "format": "float",
        "tag": "Tag",
        "supheader": "Supheader",
        "conditional": "cond_1",
        "remove_tag": True,
        "log2": True,
    }

    report_template = ReportTemplate(
        sections={
            "Standard section 1": standard_section_template,
            "Tag sample section 1": tag_sample_section_template,
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
            "sample_extraction_tag": "Tag",
            "log2_tag": "[log2]",
        },
    )
    return report_template


@pytest.fixture()
def example_table() -> pd.DataFrame:
    table = pd.DataFrame(
        {
            "Column 1": [1, 2, 3],
            "Column 2": ["A", "B", "C"],
            "Column 4": ["A", "B", "C"],
            "Tag Sample 1": ["A", "B", "C"],
            "Tag Sample 2": ["A", "B", "C"],
        }
    )
    return table


@pytest.fixture()
def standard_table_section(report_template, example_table) -> compiler.TableSection:
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
        headers={"Column 1": "Column 1", "Column 2": "Column 2"},
        header_formats={
            "Column 1": report_template.formats["header"],
            "Column 2": report_template.formats["header"],
        },
        supheader=report_template.sections["Standard section 1"]["supheader"],
        supheader_format=report_template.formats["supheader"],
        section_conditional=report_template.conditional_formats["cond_2"],
    )
    return table_section


@pytest.fixture()
def tag_sample_table_section(report_template, example_table) -> compiler.TableSection:
    table_section = compiler.TableSection(
        data=example_table[["Tag Sample 1", "Tag Sample 2"]].copy(),
        column_formats={
            "Tag Sample 1": report_template.formats["float"],
            "Tag Sample 2": report_template.formats["float"],
        },
        column_conditionals={
            "Tag Sample 1": {},
            "Tag Sample 2": {},
        },
        column_widths={
            "Tag Sample 1": 10,
            "Tag Sample 2": 10,
        },
        header_formats={
            "Tag Sample 1": report_template.formats["header"],
            "Tag Sample 2": report_template.formats["header"],
        },
        headers={"Tag Sample 1": "Sample 1", "Tag Sample 2": "Sample 2"},
        supheader="Supheader [log2]",
        supheader_format=report_template.formats["supheader"],
        section_conditional=report_template.conditional_formats["cond_1"],
    )
    return table_section


def test_eval_standard_section_columns_selects_correct_columns():
    template_section = {"columns": ["Column 1", "Column 2", "Column 3"]}
    columns = ["Column 1", "Column 2", "Column 4"]
    selected_columns = compiler.eval_standard_section_columns(template_section, columns)
    assert selected_columns == ["Column 1", "Column 2"]


def test_eval_tag_sample_section_columns_selects_correct_columns():
    template_section = {"tag": "Tag"}
    columns = ["Tag Sample 1", "Tag Sample 2", "Column 1"]
    extraction_tag = "Tag"
    selected_columns = compiler.eval_tag_sample_section_columns(
        template_section, columns, extraction_tag
    )
    assert selected_columns == ["Tag Sample 1", "Tag Sample 2"]


class TestEvalTagSampleHeaders:
    def test_with_remove_tag(self):
        template_section = {"tag": "Tag", "remove_tag": True}
        columns = ["Tag Sample 1", "Tag Sample 2"]
        expected = {"Tag Sample 1": "Sample 1", "Tag Sample 2": "Sample 2"}
        headers = compiler.eval_tag_sample_headers(columns, template_section)
        assert headers == expected

    def test_without_remove_tag(self):
        template_section = {"tag": "Tag", "remove_tag": False}
        columns = ["Tag Sample 1", "Tag Sample 2"]
        expected = {"Tag Sample 1": "Tag Sample 1", "Tag Sample 2": "Tag Sample 2"}
        headers = compiler.eval_tag_sample_headers(columns, template_section)
        assert headers == expected

    def test_without_remove_tag_and_log2_tag(self):
        template_section = {"tag": "Tag", "remove_tag": False, "log2": True}
        columns = ["Tag Sample 1", "Tag Sample 2"]
        expected = {
            "Tag Sample 1": "Tag Sample 1 [log2]",
            "Tag Sample 2": "Tag Sample 2 [log2]",
        }
        headers = compiler.eval_tag_sample_headers(columns, template_section, log2_tag="[log2]")  # fmt: skip
        assert headers == expected

    def test_with_remove_tag_and_log2_tag(self):
        template_section = {"tag": "Tag", "remove_tag": True, "log2": True}
        columns = ["Tag Sample 1", "Tag Sample 2"]
        expected = {
            "Tag Sample 1": "Sample 1",
            "Tag Sample 2": "Sample 2",
        }
        headers = compiler.eval_tag_sample_headers(columns, template_section, log2_tag="[log2]")  # fmt: skip
        assert headers == expected


class TestEvalTagSampleSupheader:
    def test_with_log2_tag(self):
        template_section = {"supheader": "Supheader", "log2": True}
        expected = "Supheader [log2]"
        headers = compiler.eval_tag_sample_supheader(template_section, log2_tag="[log2]")  # fmt: skip
        assert headers == expected

    def test_without_log2_tag(self):
        template_section = {"supheader": "Supheader", "log2": False}
        expected = "Supheader"
        headers = compiler.eval_tag_sample_supheader(template_section, log2_tag="[log2]")  # fmt: skip
        assert headers == expected


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


def test_StandardSectionCompiler(
    report_template, example_table, standard_table_section
):
    section_compiler = compiler.StandardSectionCompiler(report_template)
    template_section = report_template.sections["Standard section 1"]
    compiled_section = section_compiler.compile(template_section, example_table)
    for attr in standard_table_section.__dataclass_fields__:
        if attr == "data":
            assert compiled_section.data.equals(standard_table_section.data)
        else:
            # Include attribute name in a dictionary to get nicer error messages
            compiled_attr = {attr: getattr(compiled_section, attr)}
            expected_section_attr = {attr: getattr(standard_table_section, attr)}
            assert compiled_attr == expected_section_attr


def test_TagSampleSectionCompiler(
    report_template, example_table, tag_sample_table_section
):
    section_compiler = compiler.TagSampleSectionCompiler(report_template)
    template_section = report_template.sections["Tag sample section 1"]
    compiled_section = section_compiler.compile(template_section, example_table)
    for attr in tag_sample_table_section.__dataclass_fields__:
        if attr == "data":
            assert compiled_section.data.equals(tag_sample_table_section.data)
        else:
            # Include attribute name in a dictionary to get nicer error messages
            compiled_attr = {attr: getattr(compiled_section, attr)}
            expected_section_attr = {attr: getattr(tag_sample_table_section, attr)}
            assert compiled_attr == expected_section_attr


class TestCompileTableSection:
    def test_correctly_compiled_sections(
        self,
        report_template,
        example_table,
        standard_table_section,
        tag_sample_table_section,
    ):
        compiled_sections = compiler.compile_table_sections(report_template, example_table)  # fmt: skip
        expected_sections = [standard_table_section, tag_sample_table_section]
        for compiled_section, expected_section in zip(expected_sections, compiled_sections):  # fmt: skip
            for attr in expected_section.__dataclass_fields__:
                if attr == "data":
                    assert compiled_section.data.equals(expected_section.data)
                else:
                    # Include attribute name in a dictionary to get nicer error messages
                    compiled_attr = {attr: getattr(compiled_section, attr)}
                    expected_section_attr = {attr: getattr(expected_section, attr)}
                    assert compiled_attr == expected_section_attr

    @pytest.mark.parametrize(
        "section",
        [
            {"format": "str"},
            {"columns": ["Column 1", "Column 2"], "comparison_group": True},
            {"columns": ["Column 1", "Column 2"], "tag": "str"},
        ],
    )
    def test_invalid_sections_are_not_compiled(
        self, report_template, example_table, section
    ):
        report_template.sections = {"invalid": section}
        compiled_sections = compiler.compile_table_sections(report_template, example_table)  # fmt: skip
        assert len(compiled_sections) == 0


class TestIdentifyTemplateSectionCategory:
    def test_identify_standard_section(self):
        section_template = {"columns": []}
        section_category = compiler.identify_template_section_category(section_template)
        assert section_category == compiler.SectionCategory.STANDARD

    def test_identify_tag_sample_section(self):
        section_template = {"tag": ""}
        section_category = compiler.identify_template_section_category(section_template)
        assert section_category == compiler.SectionCategory.TAG_SAMPLE

    @pytest.mark.parametrize(
        "section_template",
        [{"comparison_group": True}, {"comparison_group": True, "tag": ""}],
    )
    def test_identify_comparison_section(self, section_template):
        section_category = compiler.identify_template_section_category(section_template)
        assert section_category == compiler.SectionCategory.COMPARISON

    @pytest.mark.parametrize(
        "section_template",
        [
            {"format": "str"},
            {"columns": [], "tag": "str"},
            {"columns": [], "comparison_group": True},
        ],
    )
    def test_identify_unknown_section(self, section_template):
        section_category = compiler.identify_template_section_category(section_template)
        assert section_category == compiler.SectionCategory.UNKNOWN
