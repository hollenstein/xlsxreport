# XlsxReport

## Introduction

XlsxReport is a Python library that simplifies the creation of well-formatted Excel reports from CSV files of quantitative mass spectrometry (MS) results. It utilizes YAML template files to specify the arrangement and formatting of the CSV content in the resulting Excel file.

With XlsxReport, generating Excel reports for mass spectrometry results from the same software or pipeline is a breeze – just create a YAML report template file once and execute a command line script to create reproducibly formatted Excel reports whenever needed.

The two main applications of XlsxReport are to create clean and uncluttered Excel files for the manual inspection of MS results, and to create Excel reports that can used as supplementary tables for publications.


## The XlsxReport report template document

To generate the formatted Excel report, XlsxReport requires an input CSV file and a report template in YAML format. The report template is used to describe the structure and formatting of the generated Excel report. This allows specifying which columns should appear in the Excel file, the order of the columns, and which columns will be grouped together into sections. The report template file allows specifying the format of headers, and applying individual formats and conditional formats to the content of each column. Moreover, it is possible to specify section supheaders that will be written above the header row into a merged cell.

It is not possible to use the report template for renaming column headers, applying calculations to column values, and for sorting rows. In general, anything that changes the data is not the scope of XlsxReport, if such a functionality is required it should be implemented in another script that can be run before XlsxReport.


### How does the report template file look like

The report template file comprises four areas named `sections`, `formats`, `conditional_formats`, and `settings`. The `sections` area is used to select and organize columns, and to specify their formatting by assigning formats and conditional formats that are defined in the `formats` and `conditional_formats` areas. For example, a format determines decimal digits or alignment, whereas conditional formats define cell appearance based on values. The `settings` area is used to define general settings like row height, whether to apply an autofilter on the header row, or if a section supheader row should be added.

Here is a simple example of a report template that is used to generate an Excel file with tree columns: "Protein ID", "Gene name", and "Spectral count". It contains only one entry, "protein_evidence", in the `sections` area. In the "protein_evidence" section three columns are selected and a default format "str" is applied to all column values. In addition, the "int" format and the "count" conditional format are specifically applied to the values of the "Spectral count" column, overriding the defaults. Finally, a supheader "Protein evidence" is defined, which will be written to the excel above the header row. Writing supheader is enabled because the `settings` area contains the entry "write_supheader: True". In the `formats` area the two formats "int" and "str" are defined that have been referenced in the "protein_evidence" section. In addition, the format specified by the "header" and "supheader" entries are applied to the header and supheader row. The `conditional_formats` area contains one conditional format called "count", which has been assigned to the "Spectral count" column in the "protein_evidence" section.

```YAML
sections:
  protein_evidence: {
    columns: ["Protein ID", "Gene name", "Spectral count"],
    column_format: {"Spectral count": "int"},
    column_conditional: {"Spectral count": "count"},
    format: "str",
    supheader: "Protein evidence",
  }

formats:
  int: {"align": "center", "num_format": "0"}
  str: {"align": "left"}
  header: {"bold": True, "align": "center", "bottom": 2}
  supheader: {"bold": True, "align": "center", "bottom": 2}

conditional_formats:
  count: {
    "type": "2_color_scale",
    "min_type": "num", "min_value": 0, "min_color": "#ffffbf",
    "max_type": "percentile", "max_value": 99.5, "max_color": "#f25540"
  }

settings:
  write_supheader: True
  add_autofilter: True
  header_height: 95
```

### Template section area - `sections`

Each entry in the `sections` area starts with a unique group name and then describes a group of columns that is written to the excel file as a block. The columns of a group can be defined manually by specifying a list of column names or automatically by defining a "tag" string that is used to extract all columns containing this string, for example "MS/MS count". In addition, each group entry contains instructions about how the columns and their headers should be formatted, if conditional formats should be applied and more. The order of groups in the config file determines the order in which the groups are written to the excel file. Columns that were already used by one group will not be used in subsequent groups.

There are currently three different categories of sections:

- In a `default section` columns are directly selected by specifying a list of column names with the `columns` parameter. The specified order of columns defines in which order the columns will be written to the Excel sheet. The parameters `tag` and `comparison_group` are not allowed in this section. The parameters `log2`, `replace_comparison_tag`, and `remove_tag` have no effect on this section type.

- The **sample group** is used to describe a block of quantitative columns. A
  sample group is defined by specifying of a column "tag", which is then used
  to automatically extract all columns that contain the "tag" as a substring.
  To extract the intensity columns "Intensity Sample_A" and "Intensity
  Sample_B" one would create a new group and specify "Intensity" as the tag.
- The **comparison group** allows defining a block of differential expression
  comparison columns. Adding the parameter "comparison_group: True" defines a
  comparison group. The columns that belong to a comparison group have a column
  name that consists of one part that describes the content of the column, for
  example "P-value" or "Fold change", and another part that describes which
  samples or experiments are compared, for example "Control vs. Condition". To
  identify comparison columns, the comparison symbol must be defined with
  the "tag" parameter, in this example the "tag" corresponds to " vs. ", and
  the strings that describe the column contents must be listed in the "columns"
  parameter, in this example ["P-value", "Fold change"]. In this example the
  comparison group would include the columns "P-value Control vs. Condition" and
  "Fold change Control vs. Condition".


### Format parameters area - `formats`
In the `formats` area the formats must be defined that are applied in the template sections. In addition, by specifying a format called "header" and "supheader" it is possible to define default formats for the header and supheader row.

Refer to the [XlsxWriter](https://xlsxwriter.readthedocs.io/format.html#format-methods-and-format-properties)
documentation for additional information which parameters can be defined for a format.

### Conditional format area - `conditional_formats`
In the `conditional_formats` area the conditional formats must be defined that are applied in the template sections.

Refer to the [XlsxWriter](https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html) documentation for additional information which parameters can be defined for a conditional format.


### Settings area - `settings`

The `settings` area is used to define general settings affecting all content that is written to the Excel sheet.

- `supheader_height: float (default: 20)`<br>
Defines the supheader row height in pixels.

- `header_height: float (default: 20)`<br>
Defines the header row height in pixels.

- `column_width: float (default: 64)`<br>
Defines default column width. This parameter is overwritten if a `width` section parameter is defined.

- `log2_tag: str (default "")`<br>
If specified this string is added as a suffix to the supheader or header of a tag section if the `log2` section parameter is defined, and a log2 transformation is applied to the column values.

- `sample_extraction_tag: str (default "")`<br>
String that is used as a substring to collect columns that contain this tag and the  sample names. These columns are then used to extract sample names. The `sample_extraction_tag` should be chosen to only select columns that contain sample names.

- `append_remaining_columns: bool (default: False)`<br>
If True, all remaining columns that are not present in any section are added to the end of the Excel sheet, and the section of appended columns is hidden.

- `write_supheader: bool (default: False)`<br>
If True, a supheader row is added above the header row.

- `evaluate_log2_transformation: bool (default: False)`<br>
If True, column values are evaluated if they appear to be already log transformed before a log2 transformation is applied.

- `remove_duplicate_columns: bool (default: True)`<br>
If True columns that are already present in a section are removed from subsequent sections.

- `add_autofilter: bool (default: True)`<br>
If True, adds an Excel auto filter to the header row.

- `freeze_cols: int (default: 1)`<br>
If a value larger than 0 is specified, freeze pane is applied in the Excel sheet. The selected row for freezing will always be the header row, the selected column is chosen based on the specified value.
