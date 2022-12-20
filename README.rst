XlsxReport
==========


Introduction
------------
XlsxReport is a Python package for automatically generating formatted excel reports from
quantitative mass spectrometry result tables. YAML config files are used to describe the
content of a result table and the format of the excel report.


Release
-------
Development is currently in early alpha and the interface is not yet stable.


Install
-------
For Windows users without Python we recommend installing the free
`Anaconda <https://www.continuum.io/downloads>`_ Python package provided by Continuum
Analytics, which already contains a large number of popular Python packages for data
science. Or get Python from the
`Python homepage <https://www.python.org/downloads/windows/>`_. XlsxReport requires
Python version 3.9 or higher.

To install XlsxReport, activate the conda environment you want to use, navigate to the
folder containing the XlsxReport files and enter the following command (don't forget to
add the dot after install):

``pip install .``


To uninstall the XlsxReport library type:

``pip uninstall xlsxreport``


After XlsxReport has been installed the local AppData directory needs to be setup and the
default configuration files need to be copied. Running the "xlsx_report_setup" script
creates a new XlsxReport folder in the local user data directory, for example
"C:/User/user_name/AppData/Local/XlsxReport" on Windows 10, and copies the default config
files there.

``xlsx_report_setup``


Run a script
------------
To generate a simple excel protein report, run the "xlsx_report" script with an input
and config file. Here is an example with the default maxquant.yaml config file.

``xlsx_report C:/proteinGroups.txt maxquant.yaml``


The script "cassiopeia_report" can be used to generate an excel protein report from the
Matrix_Export_proteinGroups.txt output of the Cassiopeia R script. In this case it is
not necessary to specify a config file, as by default the "cassiopeia.yaml" file will be
used.

``cassiopeia_report C:/Matrix_Export_proteinGroups.txt``


Planned features
----------------
- Add the option to specify column comments by providing an additional file
- Maybe add an option to specify sample names and order
    - Requires that samples are specified by user
    - _find_sample_group_columns() needs to also sort columns
