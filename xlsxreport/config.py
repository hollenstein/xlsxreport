"""This module provides a class for handling configuration for formatting excel reports.

The module contains a single class, ReportConfig, which stores the configuration of a
report. It provides methods for loading and saving the configuration to a YAML file, as
well as checking if the configuration is valid.
"""

from __future__ import annotations
from typing import Optional
import yaml


class ReportConfig:
    """Class to store the configuration of a report."""

    def __init__(
        self,
        groups: Optional[dict] = None,
        formats: Optional[dict] = None,
        settings: Optional[dict] = None,
        conditional_formats: Optional[dict] = None,
    ):
        self.groups = {} if groups is None else groups
        self.formats = {} if formats is None else formats
        self.settings = {} if settings is None else settings
        self.conditional_formats = (
            {} if conditional_formats is None else conditional_formats
        )

    @classmethod
    def load(cls, filepath) -> ReportConfig:
        """Load a report configuration from a YAML file."""
        with open(filepath, "r", encoding="utf-8") as file:
            config_data = yaml.safe_load(file)
        config = cls()
        config.formats = config_data.get("formats", {})
        config.conditional_formats = config_data.get("conditional_formats", {})
        config.groups = config_data.get("groups", {})
        config.settings = config_data.get("args", {})
        return config

    def save(self, filepath) -> None:
        """Save a report configuration to a YAML file."""
        config_data = {
            "groups": self.groups,
            "formats": self.formats,
            "conditional_formats": self.conditional_formats,
            "args": self.settings,
        }
        with open(filepath, "w", encoding="utf-8") as file:
            yaml.dump(
                config_data,
                file,
                version=(1, 2),
                sort_keys=False,
                Dumper=IndentDumper,
            )


class IndentDumper(yaml.SafeDumper):
    """Custom YAML dumper to preserve indentation."""

    def increase_indent(self, flow=False, indentless=False):
        return super(IndentDumper, self).increase_indent(flow, False)
