[![Project Status: WIP – Initial development is in progress, but there has not yet been a stable, usable release suitable for the public.](https://www.repostatus.org/badges/latest/wip.svg)](https://www.repostatus.org/#wip)

# XlsxReport

## Introduction

XlsxReport is a Python library that simplifies the creation of well-formatted Excel reports from CSV files of quantitative mass spectrometry (MS) results. It utilizes YAML template files to specify the arrangement and formatting of the CSV content in the resulting Excel file.

With XlsxReport, generating Excel reports for mass spectrometry results from the same software or pipeline is a breeze – just create a YAML report template file once and execute a command line script to create reproducibly formatted Excel reports whenever needed.

The two main applications of XlsxReport are to create clean and uncluttered Excel files for the manual inspection of MS results, and to create Excel reports that can be used as supplementary tables for publications.

## Release

XlsxReport is actively developed and currently in late alpha stage.


## Install

If you do not already have a Python installation, we recommend installing the [Anaconda distribution](https://www.continuum.io/downloads) of Continuum Analytics, which already contains a large number of popular Python packages for Data Science. Alternatively, you can also get Python from the [Python homepage](https://www.python.org/downloads/windows). XlsxReport requires Python version 3.9 or higher.

You can use pip to install XlsxReport from the distribution file with the following command:

```
pip install xlsxreport-X.Y.Z-py3-none-any.whl
```

To uninstall the XlsxReport package type:

```
pip uninstall xlsxreport
```


### Installation when using Anaconda

If you are using Anaconda, you will need to install the XlsxReport package into a conda environment. Open the Anaconda navigator, activate the conda environment you want to use, run the "CMD.exe" application to open a terminal, and then use the pip install command as described above.


### Setting up the AppData directory

After XlsxReport has been installed the local AppData directory needs to be setup and the default template files need to be copied. Running the `xlsxreport setup` script creates a new XlsxReport folder in the local user app data directory, for example "C:/User/user_name/AppData/Local/XlsxReport" on Windows 10, and copies the default template files there.

```
xlsxreport_setup
```


## Run a script

To generate a simple excel protein report, run the `xlsxreport report` script with an input and template file. Here is an example with the default maxquant.yaml template file.

```
xlsxreport proteinGroups.txt maxquant.yaml
```
