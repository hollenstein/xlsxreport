XlsxReport
==========


Introduction
------------
XlsxReport is a Python package to automatically generate formatted excel reports from
quantitative mass spectrometry result tables. YAML config files are used to describe
the content of a result table and the format of the excel report.


Release
-------
Development is currently in early alpha.


Install
-------
As the package is still in early development, it is recommended to create a virtual
environment for installing XlsxReport and testing it. 

When using the Anaconda Python distribution, you can create a new environment by
opening a terminal, entering the following command and following the instructions:

$ conda create -n xlsxreport python=3.9

To activate the environment, enter in the terminal:
$ conda activate xlsxreport

And if you want to go back to the previous environment, use the following command:
$ conda deactivate

To install XlsxReport, activate the environment you want to use, navigate to the
XlsxReport folder and enter the following:

$ pip install --editable .

To uninstall the XlsxReport package type:

$ pip uninstall xlsxreport


Run a script
------------

Running the "xlsx_report_setup" script creates a folder at User/AppData/Local/XlsxReport
for the yaml config files, and copies the default config files there.

$ xlsx_report_setup


To generate a simple excel protein report, run the "xlsx_report" script with an input and config file. Here is an example with the default maxquant.yaml config file.

$ xlsx_report C:/proteinGroups.txt maxquant.yaml


The script "cassiopeia_report" can be used to generate an excel protein report from the
Matrix_Export_proteinGroups.txt output of the Cassiopeia R script. 

$ cassiopeia_report C:/Matrix_Export_proteinGroups.txt


Planned features
----------------
- Add option to append all remaining columns (and hide them)
- Add option to specify sample order
    - Requires that samples are specified by user
    - Adapt _find_sample_group_columns() to sort columns
- Add column comments
- Add option to sort the table before writing data
