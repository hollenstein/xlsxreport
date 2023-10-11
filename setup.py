from setuptools import setup, find_packages


VERSION = "0.0.7"

packages = find_packages()
packages.append("xlsxreport.default_config")

setup(
    name="xlsxreport",
    version=VERSION,
    license="Apache v2",
    author="David M. Hollenstein",
    author_email="hollenstein.david@gmail.com",
    install_requires=["numpy", "pandas", "pyyaml", "xlsxwriter", "click", "appdirs"],
    python_requires=">=3.9",
    packages=packages,
    package_data={"xlsxreport.default_config": ["*.yaml"]},
    entry_points={
        "console_scripts": [
            "xlsxreport_setup = xlsxreport.scripts.setup_appdata_dir:cli",
            "xlsxreport = xlsxreport.scripts.report:cli",
            "cassiopeia_report = xlsxreport.scripts.cassiopeia_report:cli",
        ],
    },
    keywords=["excel", "report", "mass spectrometry", "proteomics"],
)
