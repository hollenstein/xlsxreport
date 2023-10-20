"""This module provides a class for handling configuration for formatting excel reports.

The module contains a single class, ReportTemplate, which stores the configuration of a
report. It provides methods for loading and saving the configuration to a YAML file, as
well as checking if the configuration is valid.
"""

from __future__ import annotations
from typing import Optional
import yaml


class ReportTemplate:
    """Class to store the template of a report."""

    def __init__(
        self,
        sections: Optional[dict] = None,
        formats: Optional[dict] = None,
        settings: Optional[dict] = None,
        conditional_formats: Optional[dict] = None,
    ):
        self.sections = {} if sections is None else sections
        self.formats = {} if formats is None else formats
        self.settings = {} if settings is None else settings
        self.conditional_formats = (
            {} if conditional_formats is None else conditional_formats
        )

    @classmethod
    def load(cls, filepath) -> ReportTemplate:
        """Load a report template from a YAML file."""
        with open(filepath, "r", encoding="utf-8") as file:
            template_data = yaml.safe_load(file)
        template = cls()
        template.formats = template_data.get("formats", {})
        template.conditional_formats = template_data.get("conditional_formats", {})
        template.sections = template_data.get("groups", {})
        template.settings = template_data.get("args", {})
        return template

    def save(self, filepath) -> None:
        """Save a report template to a YAML file."""
        template_data = {
            "groups": self.sections,
            "formats": self.formats,
            "conditional_formats": self.conditional_formats,
            "args": self.settings,
        }
        with open(filepath, "w", encoding="utf-8") as file:
            yaml.dump(
                template_data,
                file,
                version=(1, 2),
                sort_keys=False,
                Dumper=IndentDumper,
            )


class IndentDumper(yaml.SafeDumper):
    """Custom YAML dumper to preserve indentation."""

    def increase_indent(self, flow=False, indentless=False):
        return super(IndentDumper, self).increase_indent(flow, False)
