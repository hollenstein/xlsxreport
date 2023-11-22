# Changelog

----------------------------------------------------------------------------------------

## 0.0.8 - Refactoring and new settings

### Changed
- `groups` in the yaml config file with the setting "border: True" will now always have
  a thick border line in the written Excel file (which is a border type 2 in Excel).
- NaN entries are now always written as an empty string to the Excel file.

### Added
- Added additional parameters to the settings (`args`) section of the yaml config file,
  which allow to control behaviors that were previously applied by default.
  - `write_supheader` (default: True), determines if a supheader row will be written.
  - `evaluate_log2_transformation` (default: True), if True values are evaluated before
    applying a log2 transformation to `groups` that have the "log2: True" setting.
  - `remove_duplicate_columns` (default: True), if True columns that were already used
    in a compiled `group` are removed from subsequent groups.
  - `add_autofilter` (default: True), if True adds an Excel auto filter to header row.
  - `freeze_cols` (default: 1), if larger than 0 applies freeze pane to the Excel file.
    The selected row for freezing will always be the header row, the selected column
    corresponds to the specified value.
- Added additional settings to `groups` in the yaml config file.
  - "hide_section: True" results in sections being hidden in the Excel file.
  - A "conditional" setting can now be added to "feature" groups, which allows applying
    a conditional format to all columns of the group. 

### Removed
- (!) Removed parameters `border_weight` and `nan_symbol` from the settings (`args`)
  section of the yaml config file.

### Internal
- Changed the build config file to `pyproject.toml`.
- Added extensive unit testing.
- Added an integration test for generating a formatted Excel file.

----------------------------------------------------------------------------------------

## 0.0.7 - Fix comparison group issue

### Fixes
  - Fixes a mix up of columns in comparison groups that was caused when an experiment 
    comparison was the exact substring of another, for example "exp1 vs exp2" and
    "exp3exp1 vs exp2".

----------------------------------------------------------------------------------------

## 0.0.6 - Fix missing supheader

### Fixes
  - Supheader not being written when a block contains only one column.

----------------------------------------------------------------------------------------

## 0.0.5 - Improvements for MsReport report generation

### Changed
- (!) Renamed console script "xlsx_report" to "xlsxreport"
- (!) Renamed console script "xlsx_report_setup" to "xlsxreport_setup"
- Columns retrieved for "Sample group" blocks now must contain the group tag and a
  sample name. Columns not containing a sample name are ignored. For example, using the
  tag "Intensity" will no longer include columns such as "Intensity" or
  "Intensity total".
- Renamed the "qtable_proteins.yaml" config file to "msreport_lfq_protein.yaml"
- Updated "msreport_lfq_protein.yaml" config file
  - Changed columns in the "protein_features" block
  - Changed comparison group tag from "logFC" to "Ratio [log2]"
  - Added new format for "Ratio [log2]"
  - Changed formatting of the quantified_events block  
- Updated console ouput for the "xlsxreport_setup" console script.

----------------------------------------------------------------------------------------

## 0.0.4 - Group and hide remaining columns

- Changes behavior of the "append_remaining_columns" option. Now the
  unspecified columns that are added to the end of the excel sheet
  are grouped and hidden.

----------------------------------------------------------------------------------------

## 0.0.3 - Installation fix

- Fixes issues of missing .yaml config files for installation.

----------------------------------------------------------------------------------------

## 0.0.2 - Adds remaining columns to report

- Added option to add all unspecified columns to the end of the excel sheet.
- Added documentation of the config file format.
- The config file argument "remove_tag" does not affect comparison groups
  anymore, as the sample comparison string is now always removed from the
  header and added to the supheader.
- Minor changes to the default config files.
- The xlsx_report_setup script now prints its progress to the console.

----------------------------------------------------------------------------------------

## 0.0.1 - First functional version

- Initial unstable version of XlsxWriter
