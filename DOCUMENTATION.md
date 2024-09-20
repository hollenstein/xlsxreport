# XlsxReport

The documentation for XlsxReport is still work in progress. This file provides an overview of how XlsxReport works and a detailed description of the table template and its formatting options.

## Table of contents
- [How does **XlsxReport** work and what can it do?](#how-does-xlsxreport-work-and-what-can-it-do)
  - [What **XlsxReport** can't do](#what-xlsxreport-cant-do)
- [How does the table template look like](#how-does-the-table-template-look-like)
- [The `sections` area of the table template](#the-sections-area-of-the-table-template)
  - [1) The standard template section](#1-the-standard-template-section)
  - [2) Tag template section](#2-tag-template-section)
  - [3) Label tag template section](#3-label-tag-template-section)
  - [4) Comparison template section](#4-comparison-template-section)
  - [5) Common optional template section parameters](#5-common-optional-template-section-parameters)
- [The `formats` area of the table template](#the-formats-area-of-the-table-template)
- [The `conditional_formats` are of the table template](#the-conditional_formats-are-of-the-table-template)
- [The `settings` area of the table template](#the-settings-area-of-the-table-template)


## How does XlsxReport work and what can it do?

To generate the formatted Excel report, XlsxReport requires tabular data, for example a CSV file, and a table template, a YAML file that contains instructions for the structure and formatting of the generated Excel report. The table template allows specifying which columns appear in the Excel file, the order of the columns, and which columns will be grouped together into `sections`. Furthermore, the format of headers can be specified, and individual formats and conditional formats can be applied to the content of each column. Moreover, it is possible to specify `section` supheaders that will be written above the header row into a merged cell.

> _**Note**: The term `section` or `template section` refers to a group of columns that are defined in the table template as a unit, and that are written to the Excel sheet as a block, or as a "section"._


### What XlsxReport can't do

In general, changing values, i.e. the content of columns, is beyond the scope of XlsxReport. Specifically, it is not possible to use the table template for applying calculations to columns or for filtering and sorting rows. If such functionality is required it should be implemented in another script that can be run before XlsxReport. The only exception to this is the possibility to apply a log2 transformation on the values of a `tag section` or `label tag section`. 

In addition, directly renaming column headers in the `standard section` is not yet supported, but is planned for a future release. 


## How does the table template look like

The table template comprises four areas named `sections`, `formats`, `conditional_formats`, and `settings`. Each of these is encoded as a mapping in the YAML file:

```YAML
sections: {}
formats: {}
conditional_formats: {}
settings: {}
```

The `sections` area is used to select and organize columns, and to specify their formatting by assigning formats and conditional formats that are defined in the `formats` and `conditional_formats` areas. For example, a format can specify the number of decimal digits that are displayed or the text alignment, whereas conditional formats define cell appearance based on values. Basically, formats and conditional formats are defined by using the same parameters as in Excel. Finally, the `settings` area of the table template is used to define general settings, such as row height, whether to apply an autofilter on the header row, or if an additional row for `section` supheaders should be added.

Here is a basic example of a table template that generates an Excel file with three columns "Protein IDs", "Gene names", and "Unique peptides".

```YAML
sections:
  protein_evidence:
    columns: ["Protein IDs", "Gene names", "Unique peptides"]
    column_format: {"Unique peptides": "int"}
    column_conditional_format: {"Unique peptides": "count"}
    format: "str"
    supheader: "Protein evidence"

formats:
  int: {"align": "center", "num_format": "0"}
  str: {"align": "left"}
  header: {"bold": True, "align": "center", "bottom": 2}
  supheader: {"bold": True, "align": "center", "bottom": 2}

conditional_formats:
  count:
    "type": "2_color_scale"
    "min_type": "num"
    "min_value": 0
    "min_color": "#ffffbf"
    "max_type": "percentile"
    "max_value": 99.5
    "max_color": "#f25540"

settings:
  write_supheader: True
```

This table template example contains only one `template section` with the internal name "protein_evidence". The `columns` keyword is used to select which columns will appear in the Excel report. Using the argument `format` the default format "str" is applied to all columns within the `template section`. In addition, the "int" format and the "count" conditional format are specifically applied to the values of the "Unique peptides" column, overriding the defaults. Finally, a supheader "Protein evidence" is defined, which will be written to the Excel sheet above the header row. Writing a supheader row must be specifically enabled, which is done in the `settings` area with the parameter "write_supheader: True". In the `formats` area the two formats "int" and "str" are defined that have been referenced in the "protein_evidence" `template section`. In addition, the format specified with the name "header" and "supheader" are special formats that are always applied to the header and supheader rows. The `conditional_formats` area contains the description of the conditional format called "count", which has been assigned to the "Unique peptides" column.


## The `sections` area of the table template

Each entry in the `sections` area is defined by a unique internal name and contains parameters that are used to select one or multiple columns, and parameters that define how the columns are formatted in the Excel report. There are currently three different types of template `sections`, each provides a different way how the columns for the section are selected. In addition, the parameters specified in a `template section` describe how the column values and headers will be formatted, if conditional formats are applied, and other settings. The order of `template sections` in the table template determines the order in which the sections, and thus the columns, are written to the Excel file.


### 1) The standard template section

When using a `standard template section`, you can directly select the columns that should be written to the Excel sheet by specifying a list of column names. This is useful when column names are constant in your output and don’t change between experiments, such as a "Protein IDs" column. Formats and conditional formats can be applied to the whole `section` or to individual columns.

#### Required and optional section parameters

- `columns`<br>
A list of column names that should be written to the Excel sheet. The order of the columns in the sequence defines the order in which the columns will be written to the Excel sheet. Columns that are not present in the input table are ignored.
  - Type: `sequence[string]`
  - Is required

For additional optional parameters refer to the [common optional template section parameters](#5-common-optional-template-section-parameters).


### 2) Tag template section
You can use the `tag template section` to select a group of columns based on a regular expression pattern that is matched against all column names. This is useful when the column names are not constant between multiple tables, but contain a constant common and a variable part. For example, when you have columns named "Intensity Sample_1", "Intensity Sample_2", and so on, you can use a `tag template section` to select all columns that start with "Intensity".

> _**TIP:** Use the regular expression anchors `^` and `$` to specify if the column needs to begin or end with the tag. To exclude columns that are an exact match to the specified tag, add a `.` in front or after the tag to indicate that at least one additional character must be present._

#### Required and optional section parameters

- `tag`<br>
A regular expression pattern that is matched against all column names. Columns that match the pattern are selected for the section. Columns are written to the Excel sheet in the order they appear in the input table.
  - Type: `string`
  - Is required

- `remove_tag`<br>
If True, the matched regular expression pattern of the `tag` is removed from the column name in the Excel sheet. Removing the tag is useful in combination with adding a `supheader` to the section, as in this way the removed tag can be displayed only in the supheader row.
  - Type: `bool`
  - Default: `False`
  - Is optional

- `log2`<br>
Use this parameter to apply a log2 transformation to the values of all selected columns in the `tag template section`. The global setting parameters `log2_tag` and `evaluate_log2_transformation` affect the behavior of the `tag template section` when the `log2` parameter is used. Refer to the description of the [table template settings](#the-settings-area-of-the-table-template) for more information.
  - Type: `bool`
  - Default: `False`
  - Is optional

For additional optional parameters refer to the [common optional template section parameters](#5-common-optional-template-section-parameters).


#### How does the `template tag section` look like in practice?

Let's assume your have the following table template, containing a `tag sample section` with the name "intensities":

```YAML
sections:
  intensities: {tag: "^Intensity."}
```

and a CSV file with the following columns

| Protein IDs | Intensity sample_1 | Intensity sample_2 | Mean Intensity |
| ----------- | ------------------ | ------------------ | -------------- |
| P40238      | 1,000              | 2,000              | 1,500          |

 When generating a report, XlsxReport selects columns that match the regular expression pattern `^Intensity.`, which results in the selection of the columns "Intensity sample_1" and "Intensity sample_2". The specified pattern requires that a column starts with "Intensity" and that "Intensity" is followed by an additional character. This pattern does not match the column "Mean Intensity", which is therefore not included in the `section`.


### 3) Label tag template section

The `label tag template section` is an extension of the `tag template section` that enables more precise control over the selection of columns. While you only specify the constant part of a column name with the `tag` parameter in the `tag template section`, the `label tag template section` allows you to also specify the variable part of the columns that should be included in the section. This is achieved with the `labels` parameter, which is a list representing the variable part of the columns that should be included in the `section`. For a column to be selected, the column name must exactly contain the constant part specified in the `tag` parameter and one of the variable parts specified in the `labels` parameter.

As the `label tag template section` requires preexisting knowledge of the variable part of column names, such as sample names that are expected to change between datasets, it is typically not used in `table templates` that are intended as a general template for a specific data analysis pipeline. However, it is very useful to dynamically adjust a table template to a specific dataset. For example, when you want to create a report but only select a subset of the columns that are matched by the `tag` parameter.

#### Required and optional section parameters

- `tag`<br>
A regular expression pattern that is matched against all column names. Columns that match the pattern are selected for the section. Columns are written to the Excel sheet in the order they appear in the input table.
  - Type: `string`
  - Is required

- `labels`<br>
A sequence of strings that represent the variable part of the column names that should be included in the section.
  - Type: `sequence[string]`
  - Is required

- `remove_tag`<br>
If True, the matched regular expression pattern of the `tag` is removed from the column name in the Excel sheet. Removing the tag is useful in combination with adding a `supheader` to the section, as in this way the removed tag can be displayed only in the supheader row.
  - Type: `bool`
  - Default: `False`
  - Is optional

- `log2`<br>
Use this parameter to apply a log2 transformation to the values of all selected columns in the `tag template section`. The global setting parameters `log2_tag` and `evaluate_log2_transformation` affect the behavior of the `tag template section` when the `log2` parameter is used. Refer to the description of the [table template settings](#the-settings-area-of-the-table-template) for more information.
  - Type: `bool`
  - Default: `False`
  - Is optional

For additional optional parameters refer to the [common optional template section parameters](#5-common-optional-template-section-parameters).


### 4) Comparison template section

The `comparison template section` allows you to select and group columns that represent pair-wise comparisons of conditions. Examples of such columns include statistical comparisons of differences between two experimental conditions or the ratio between the mean intensities of those conditions. To be included in a `comparison section`, column names must follow a specific logic:

1. **Condition Description**: Column names must contain a part that describes the conditions being compared, such as "Control vs. Condition". It is crucial to have a consistent symbol between the two conditions to identify the comparison columns. This symbol is defined by the `tag` parameter. In the example, the `tag` is " vs. ".
2. **Comparison Type**: Column names must also include a part that describes the type of comparison, such as "P-value" or "Ratio". Multiple types of comparisons can be included in a `comparison section`, and the substrings that identify the comparison type are listed in the `columns` parameter.


#### Required and optional section parameters

- `tag`<br>
A string that corresponds to the comparison symbol between two conditions. The `tag` is used to pre-select columns that might belong to the `comparison section`.
  - Type: `string`
  - Is required

- `columns`<br>
A sequence of strings that correspond to the substrings that identify the comparison type in the column names.
  - Type: `sequence[string]`
  - Is required

- `replace_comparison_tag`<br>
Optional parameter that allows you to replace the comparison tag with a different string. This can be useful if you want to change the comparison tag in the Excel sheet to make it more readable.
  - Type: `string`
  - Is optional

- `remove_tag`<br>
If True, the condition comparison string is removed from the columns, leaving only the comparison type that was specified with the `columns` parameter. This option is useful in combination with adding a `supheader` to the section, as in this way the condition comparison string is only displayed in the supheader row.
  - Type: `bool`
  - Default: `False`
  - Is optional

For additional optional parameters refer to the [common optional template section parameters](#5-common-optional-template-section-parameters).

> _**NOTE:** Several optional parameters work slightly different in this `section` type The parameters `column_format` and `column_conditional_format` are used to apply formats and conditional formats to columns that contain a specific comparison type, for example "P-value" or "Ratio". The column names that are specified for these parameters therefore need to correspond to entries from the `columns` parameter. The `supheader` parameter has no effect, as the supheader is automatically generated from the column names and corresponds to the conditions that are compared._


#### How does the `comparison template section` work in practice?

Let's look at an example to illustrate how the `comparison template section` works. Assume you have the following columns in your input table:

- "Ratio Control vs. Condition"
- "Ratio Control vs. Another condition"
- "P-value Control vs. Condition"
- "P-value Control vs. Another condition"
- "Intensity Control vs. Condition"
- "Intensity Control vs. Another condition"

And the following table template:

```YAML
sections:
  statistical_comparison:
    tag: " vs. "
    columns: ["P-value", "Ratio"]
```

When generating a report, XlsxReport first collects all comparison columns, then groups them according to the conditions that are compared. All columns that compare the same two conditions are then used to write a separate `section` to the Excel sheet:

In the example the first section contains the columns:

- "Ratio Control vs. Condition"
- "P-value Control vs. Condition"

And the second section contains the columns:

- "Ratio Control vs. Another condition"
- "P-value Control vs. Another condition"

The columns "Intensity Control vs. Condition" and "Intensity Control vs. Another condition" are not included in any `section`, since "Intensity" was not listed in the `columns` parameter of the `comparison section`.


### 5) Common optional template section parameters

The following optional parameters can be used in all `template sections` types to specify formatting and other settings:

- `format`<br>
The default format that is applied to all columns in the `section`. The format must be defined in the `formats` area of the table template.
  - Type: `string`

- `column_format`<br>
A mapping that specifies formats that are applied to individual columns in the `section`. The column format overrides the default format. The keys are column names, and the values are format names that are defined in the `formats` area of the table template.
  - Type: `mapping[string, string]`

- `conditional_format`<br>
The name of the conditional format that is applied to the values of all columns in the `section`. The conditional format must be defined in the `conditional_formats` area of the table template.
  - Type: `string`

- `column_conditional_format`<br>
A mapping that specifies the conditional format that is applied to the values of individual columns in the `section`. The keys are column names, and the values are conditional format names that are defined in the `conditional_formats` area of the table template.
  - Type: `mapping[string, string]`

- `header_format`<br>
Allows to specify additional formatting properties that are applied to the header format of the `section`. The specified formatting properties are added to the default "header" format that can be defined in the `formats` area of the table template. For more information about how to define formatting properties refer to the [documentation of the formats area](#the-formats-area-of-the-table-template).
  - Type: `mapping[string, mapping]`

- `supheader`<br>
A string that is written to the Excel sheet above the header row of the `section`. The `supheader` is written to a merged cell that spans all columns of the `section`. The `supheader` is only written if the global setting `write_supheader` is set to `True`.
  - Type: `string `

- `supheader_format`<br>
Allows to specify additional formatting properties that are applied to the super header format of the `section`. The specified formatting properties are added to the default "supheader" format that can be defined in the `formats` area of the table template. For more information about how to define formatting properties refer to the [documentation of the formats area](#the-formats-area-of-the-table-template).
  - Type: `mapping[string, mapping]`

- `width`<br>
Defines the column widths in pixels. The width is applied to all columns in the `section` and overwrites the default column width that is defined in the `settings` area of the table template.
  - Type: `float`

- `border`<br>
If set to True, a thick border line is added to the left and right side of the section.
  - Type: `boolean`
  - Default: `False`


## The `formats` area of the table template

In the `formats` area the formats must be defined that are applied in the `sections` area of the table template. In addition, by specifying a format called "header" and "supheader" it is possible to define the default formats for the header and supheader row.

Refer to the [XlsxWriter](https://xlsxwriter.readthedocs.io/format.html#format-methods-and-format-properties)
documentation for additional information which parameters can be defined for a format. Note that entries of the **Property** column from the documentation correspond to the keys that can be defined in the format mapping.

Here is an example of a `formats` area that defines the formats "int", "float", "str", "header", and "supheader":

```YAML
formats:
  int: {"align": "center", "num_format": "0"}
  float: {"align": "center", "num_format": "0.00"}
  str: {"align": "left", "num_format": "0"}
  header: {
    "bold": True,
    "align": "center",
    "valign": "vcenter",
    "bottom": 2,
    "top": 2,
    "text_wrap": True
  }
  supheader: {
    "bold": True,
    "align": "center",
    "valign": "vcenter",
    "bottom": 2,
    "left": 2,
    "right": 2,
    "text_wrap": True
  }
```


## The `conditional_formats` are of the table template

In the `conditional_formats` area the conditional formats must be defined that are applied in the template sections.

The type of conditional format needs to be defined with the `type` parameter. Currently only the types `2_color_scale`, `3_color_scale`, and `data_bar` are supported. In addition, the formatting parameters corresponding to the selected type must be defined.

Refer to the [XlsxWriter](https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html) documentation for additional information which parameters can be defined for different conditional format types.

Here is an example of a `conditional_formats` area that defines a 3-color scale conditional format called "intensity":

```YAML
conditional_formats:
  intensity: {
    "type": "3_color_scale",
    "min_type": "min", "min_color": "#2c7bb6",
    "mid_type": "percentile", "mid_value": 50, "mid_color": "#ffffbf",
    "max_type": "max", "max_color": "#f25540"
  }
```


## The `settings` area of the table template

The `settings` area is used to define general settings affecting all content that is written to the Excel sheet.

- `supheader_height`<br>
Defines the supheader row height in pixels.
  - Type: `float`
  - Default: `20`

- `header_height`<br>
Defines the header row height in pixels.
  - Type: `float`
  - Default: `20`

- `column_width`<br>
Defines the default column widths in pixels. This parameter is overwritten by the `width` parameter specified in the `template sections`.
  - Type: `float`
  - Default: `64`

- `log2_tag`<br>
If specified, this tag is added as a suffix to the column headers of a `tag template section` or a `label tag template section` to indicate that a log2 transformation has been applied. The `log2_tag` is only added if the section parameter `log2` is set to `True`. If the section parameter `remove_tag` is set to True, the `log2_tag` is added to the section supheader instead of the column headers.
  - Type: `str`
  - Default ""

- `append_remaining_columns`<br>
If True, all remaining columns that are not present in any section are added to the end of the Excel sheet, and the section of appended columns is hidden.
  - Type: `bool`
  - Default: `False`

- `write_supheader`<br>
If True, a supheader row is added above the header row.
  - Type: `bool`
  - Default: `False`

- `evaluate_log2_transformation`<br>
**Use this setting with caution!** If True, column values are evaluated if they appear to be already log transformed before a log2 transformation is applied. Assumes that values are log transformed if all values in a column are smaller or equal to 64. Intensities values (and intensity peak areas) reported by tandem mass spectrometry typically range from 10^1 to 10^12. To reach log2 transformed values greater than 64, intensities would need to be higher than 10^19, which seems to be very unlikely to be ever encountered.
  - Type: `bool`
  - Default: `False`

- `remove_duplicate_columns`<br>
If True, columns that are already present in a section are removed from subsequent sections. This option guarantees that columns are not duplicated in the Excel sheet.
  - Type: `bool`
  - Default: `True`

- `add_autofilter`<br>
If True, adds an Excel auto filter to the header row.
  - Type: `bool`
  - Default: `True`

- `freeze_cols`<br>
If a value larger than 0 is specified, freeze pane is applied in the Excel sheet. The selected row for freezing will always be the header row, the position of the selected column is defined by specified value.
  - Type: `int`
  - Default: `1`
