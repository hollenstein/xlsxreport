YAML report configuration file
==============================

General
-------

Note that the XlsxReport is still at an early development stage and the
specifications of the YAML configuration file might change.

The config file is used to describe the structure and formatting of the
generated excel report. The config file contains four main sections: "groups",
"formats", "conditional_formats", and "args". The "groups" section allows
specifying blocks of columns that are written to the excel sheet and how these
columns are formatted. The "formats" and "conditional_formats" sections are
used to define normal and conditional formats that are assigned in the "groups"
section. In the "args" section general settings are defined.


**Example of a simple config file with one feature and one sample group**

- Note that all key words in the config file are written in lower case, and
  multiple words are separated by a underscore.
- Each entry within a section consists of a key, value pair that is written as
  "key: value". The values in the "args" section are either a string, number or
  boolean; whereas in the other sections each value itself is a dictionary that
  contains key, value pairs.
- Multiple entries within a section are not separated by a comma.

::

  groups:
    features: {
      format: "str",
      columns: [
        "Protein IDs",
        "Protein names",
      ],
    }
    intensity: {
      tag: "Intensity",
      format: "float",
      conditional: "intensity",
      supheader: "Intensity",
      remove_tag: True,
      border: True,
    }

  formats:
    float: {"align": "center", "num_format": "0.00"}
    str: {"align": "left", "num_format": "0"}
    header: {"bold": True, "align": "center", "bottom": 2}
    supheader: {"bold": True, "align": "center", "bottom": 2}

  conditional_formats:
    intensity: {
      "type": "3_color_scale",
      "min_type": "min", "min_color": "#2c7bb6",
      "mid_type": "percentile", "mid_value": 50, "mid_color": "#ffffbf",
      "max_type": "max", "max_color": "#f25540"
    }

  args:
    column_width: 45



Group section - "``groups``"
----------------------------

Each entry in the "groups" section starts with a unique group name and then
describes a group of columns that is written to the excel file as a block. The
columns of a group can be defined manually by specifying a list of column names
or automatically by defining a "tag" string that is used to extract all columns
containing this string, for example "MS/MS count". In addition, each group
entry contains instructions about how the columns and their headers should be
formatted, if conditional formats should be applied and more. The order of
groups in the config file determines the order in which the groups are written
to the excel file. Columns that were already used by one group will not be used
in subsequent groups.

There are three different types of groups:

- The **feature group** requires that columns are directly specified by defining
  a list of column names.
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



Group parameters
~~~~~~~~~~~~~~~~

``format: "format_name"``

- References a format in the "formats" section of the config file.

``columns: ["Column name 1", "Column name 2", ...]``

- A list of column names that will be written to the excel sheet.

``width: integer (in pixel)``

- Optional, changes the default column width for this group.

``column_format: {"Column name 1": "format_name", ...}``

- Optional, allows overriding the group format for individual columns.
  References a format in the "formats" section of the config file.

``column_conditional: {"Column name 1": "conditional_format_name", ...}``

- Optional, allows specifying a conditional format for individual columns.
  References a conditional format in the "conditional_formats" section of the
  config file. The keys in "conditional_formats" have to correspond to values
  present in the "column" parameter, also for comparison groups.

``border: bool (True or False)``

- Optional, if True a border is added on the left side of the first group column
  and on the right side of the last group column. Default value is False.

``header_format: {"Argument 1": "Value 1", ...}``

- Allows addition or modification of format arguments to the default header
  format.

``supheader: string``

- Defines a group super header. If specified, the cells in the first row, above
  the group columns are merged and the supheader text is added. Has no effect
  in comparison groups.

``supheader_format: {"Argument 1": "Value 1", ...}``

- Allows the addition or modification of format arguments to the default
  supheader format.

``tag: string``

- If specified, turns a group into a sample group. Is used to collect all
  columns that contain the "tag" as a substring, these columns are then written
  to the excel sheet.

``log2: bool (True or False)``

- If True, applies a log2 transformation to the column entries. Tries to guess
  if the values were already log transformed, in which case the transformation
  is not applied. Only intended for usage in sample groups and should only be
  applied to intensity columns, not e.g. to spectral count data. Default value
  is False.

``conditional: string``

- Applies a conditional format to all columns of a sample group. References a
  conditional format in the "conditional_formats" section of the config file.
  Has not effect on feature or comparison groups.

``comparison_group: bool (True or False)``

- If true, turns a sample group into a comparison group. Supheader is
  automatically generated. Default value is False. MISSING EXPLANATION

``replace_comparison_tag: string``

- If specified, replaces the "tag" string in the supheader with. This argument
  only affects comparison groups.

``remove_tag: bool (True or False)``

- If True, removes the specified "tag" string from the column headers. Intended
  to be used together with a supheader. This argument does not affect
  comparison groups. Default value is False.



Format section - "``formats``"
------------------------------

In the formats section all formats are defined that can be used in the groups
section. The formats section must at least define a "header" and "supheader"
format.

Refer to the `XlsxWriter <https://xlsxwriter.readthedocs.io/format.html/>`_
documentation (Section: Format methods and Format properties) for further
information.



Conditional format section - "``conditional_formats``"
------------------------------------------------------

In the conditional format section all conditional formats are defined that can
be used in the groups section.

Refer to the `XlsxWriter <https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html/>`_
documentation for further information.



Arguments section - "``args``"
------------------------------

The section "args" is used to set some general options.

``border_weight: integer (pixel)``:

- Specifies the border weight when using the group argument "border". 

``supheader_height: integer (in pixel)``:

- Specify the supheader row height.

``header_height: integer (in pixel)``:

- Specify the header row height.

``column_width: integer (in pixel)``:

- Specify the default column width.

``nan_symbol: string``:

- Replaces nan values in numeric columns with this string.

``log2_tag: string``:

- If specified this string is added as a suffix to the supheader or header if
  sample columns are log2 transformed.

``sample_extraction_tag: string``: "Intensity"

- String that is used as a substring to collect columns that contain this tag
  and the sample names. These columns are then used to extract sample names.
  The "sample_extraction_tag" should only be present in columns that also
  contain sample names.

``append_remaining_columns: bool (True or False)``:

- If True, then all remaining columns that were not added by any of
  the "groups" are appended to the end of the excel sheet.
