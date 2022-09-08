from setuptools import setup, find_packages


VERSION = "0.0.2"
setup(
    name="xlsxreport",
    version=VERSION,
    license="Apache v2",
    author="David M. Hollenstein",
    author_email="hollenstein.david@gmail.com",
    install_requires=["numpy", "pandas", "pyyaml", "xlsxwriter", "click", "appdirs"],
    python_requires=">=3.9",
    packages=find_packages(),
    entry_points={
        "console_scripts": [
            "xlsx_report_setup = xlsxreport.scripts.setup_appdata_dir:cli",
            "xlsx_report = xlsxreport.scripts.report:cli",
            "cassiopeia_report = xlsxreport.scripts.cassiopeia_report:cli",
        ],
    },
    keywords=["excel", "report", "mass spectrometry", "proteomics"],
)
