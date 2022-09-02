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
opening a terminal and entering the following command:

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
