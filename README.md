# XlsxReport
[![Project Status: Active – The project has reached a stable, usable state and is being actively developed.](https://www.repostatus.org/badges/latest/active.svg)](https://www.repostatus.org/#active)
![Python Version from PEP 621 TOML](https://img.shields.io/python/required-version-toml?tomlFilePath=https%3A%2F%2Fraw.githubusercontent.com%2Fhollenstein%2Fprofasta%2Fmain%2Fpyproject.toml)
[![pypi](https://img.shields.io/pypi/v/xlsxreport)](https://pypi.org/project/xlsxreport)
[![Run pytest](https://github.com/hollenstein/xlsxreport/actions/workflows/run-pytest.yml/badge.svg?branch=main)](https://github.com/hollenstein/xlsxreport/actions/workflows/run-pytest.yml)

**XlsxReport** is a Python library that automates the creation of formatted Excel reports from tabular data.


## Table of Contents

- [What is XlsxReport?](#what-is-xlsxreport)
- [Getting Started with a simple example](#getting-started-with-a-simple-example)
- [Installation](#installation)
    - [Setting up the application data directory](#setting-up-the-application-data-directory)
    - [Installation when using Anaconda](#installation-when-using-anaconda)
- [Additional project information](#additional-project-information)
    - [Documenation](#documenation)
    - [Upcoming features and work in progress](#upcoming-features-and-work-in-progress)
    - [Do you have feedback or need help?](#do-you-have-feedback-or-need-help)


## What is XlsxReport?

Well-formatted Excel reports are important for presenting and sharing data in a clear and structured manner with collaborators, in publications, and for the manual inspection of results. However, creating these reports manually is time-consuming, tedious, and has to be repeated for every new dataset and analysis. XlsxReport was developed to streamline the process of turning tabular data into formatted Excel reports. By automating this task, XlsxReport allows the creation of consistent, publication-ready Excel reports with minimal effort.

XlsxReport uses YAML template files to define the content, structure, and formatting of the generated Excel reports. The library provides a command line interface and a Python API, allowing users to create Excel reports by applying table templates to tabular data. Although XlsxReport has been developed for quantitative mass spectrometry data, its versatile design makes it suitable for any type of tabular data.

XlsxReport is actively developed as part of the computational toolbox for the [Mass Spectrometry Facility](https://www.maxperutzlabs.ac.at/research/facilities/mass-spectrometry-facility) at the Max Perutz Labs (University of Vienna).

## Getting Started with a simple example

With XlsxReport, generating reproducibly formatted Excel reports from your data analysis pipeline is a breeze - simply create a YAML table template once and execute a single command on the command line to create Excel reports whenever needed.

Give it a try by using the provided example files in the `examples` directory. The `examples` directory contains a "proteinGroups.txt" file from MaxQuant, which can be turned into a formatted Excel report with the included default table template file "maxquant.yaml".

After installing XlsxReport and setting up the application data directory as described below, you can create an Excel report by running the following command in the command line:

```shell
xlsxreport compile examples/proteinGroups.txt maxquant.yaml
```

This command will create an Excel file named "proteinGroups.report.xlsx" in the same directory as the input file. The Excel file contains the data from the input file formatted according to the instructions in the table template.

You can achieve the same result using the Python API with the following code:

```python
import pandas as pd
import xlsxreport

template_path = xlsxreport.get_template_path("maxquant.yaml")
template = xlsxreport.TableTemplate.load(template_path)
table = pd.read_csv("./examples/proteinGroups.txt", sep="\t")
with xlsxreport.ReportBuilder("./examples/proteinGroups.report.xlsx") as builder:
    builder.add_report_table(table, template, tab_name="Report")
```

> _**NOTE:** The `xlsxreport compile` command and the `xlsxreport.get_template_path` Python function will initially verify if a valid file path for the table template is provided. If the table template file is not found, the application data directory will be searched. This feature allows you to store your default table templates in the application data directory and use them without specifying the full path._


## Installation

If you do not already have a Python installation, we recommend installing the [Anaconda distribution](https://www.anaconda.com/download) or [Miniconda](https://docs.anaconda.com/free/miniconda/index.html) distribution from Continuum Analytics, which already contains a large number of popular Python packages for Data Science. Alternatively, you can also get Python from the [Python homepage](https://www.python.org/downloads/windows). Note that XlsxReport requires Python version 3.9 or higher.

The following command will install the latest version of XlsxReport and its dependencies from PyPi, the Python Packaging Index:

```shell
pip install xlsxreport
```

To uninstall the XlsxReport library use:

```shell
pip uninstall xlsxreport
```


### Setting up the application data directory

After XlsxReport has been installed you should create the local application data directory, which enables more convenient access to your default table templates. Running the following command creates a new XlsxReport folder in the local user application data directory, for example "C:/User/user_name/AppData/Local/XlsxReport" on Windows 10, and copies the default table templates that are included with XlsxReport:

```shell
xlsxreport appdir --setup
```

To view the path to the application data directory, you can run the following command:

```shell
xlsxreport appdir
```

Including the `--reveal` flag will open the application data directory in the file explorer:

```shell
xlsxreport appdir --reveal
```


### Installation when using Anaconda

To install the XlsxReport package using Anaconda, you need to either activate a custom conda environment or install it into the default base environment. Open the Anaconda Navigator, activate the desired conda environment or use the base environment, and then open a command line by running the "CMD.exe" application. Finally, use the `pip install` command as previously before.


## Additional project information


### Documenation

The documentation of XlsxReport is work in progress. In the meantime, you can find a detailed description of the table template and its formatting options in the [DOCUMENATION.md](https://github.com/hollenstein/xlsxreport/blob/main/DOCUMENTATION.md) file on the GitHub repository.

The Python API is currently documented only in the source code. The stable public API comprises the functions and classes that are directly present in the `xlsxreport` namespace, please refer to the `xlsxreport/__init__.py` file for more information

For more information about the **command line interface**, you can run the following command:

```shell
xlsxreport --help
```

To get help for a specific command (`appdir`, `compile`, or `validate`), you can run:

```shell
xlsxreport <command> --help
```

You can find a comprehensive record of changes in the [CHANGELOG.md](https://github.com/hollenstein/xlsxreport/blob/main/CHANGELOG.md) file.


### Upcoming features and work in progress

The library has reached a stable state and we are currently working on **extending the documentation** and adding **minor feature enhancements**. In addition, we are planning to release a **simple GUI** for creating Excel reports that provides the same functionality as the command line interface and lowers the barrier for users who are not comfortable with using the command line.

### Do you have feedback or need help?

If you have any feature requests, suggestions, or bug reports, please feel free to open an issue on the [GitHub issue tracker](https://github.com/hollenstein/xlsxreport/issues).

You don't know how to use the library, or you have a question? Please feel free to contact us via email or on GitHub. We are happy to help you get started with XlsxReport and answer any questions you might have.
