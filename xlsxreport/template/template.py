"""Module for storing, loading and manipulating report templates.

The `ReportTemplate` class is a Python representation of a YAML report template file and
contains the configuration instructions for compiling a table into a formatted Excel.
The `ReportTemplate` class provides methods for loading a report template from a YAML
file, and saving the template to a YAML file.
"""

from __future__ import annotations
from typing import Optional

import yaml

from xlsxreport.validate import (
    validate_document_entry_types,
    validate_template_file_integrity,
)
from xlsxreport.template.sections import ReportTemplateSections
from xlsxreport.template.settings import ReportTemplateSettings
from xlsxreport.template.formats import ReportTemplateFormats


class ReportTemplate:
    """Class to store the template of a report.

    # Mention that

    Attributes:
        sections: A dictionary of sections in the report template. The keys are the
            names of the template sections, the values are dictionaries with the section
            parameters.
        formats: A dictionary of formats in the report template. The keys are the names
            of the formats, the values are dictionaries with the format parameters.
        conditional_formats: A dictionary of conditional formats in the report template.
            The keys are the names of the conditional formats, the values are
            dictionaries with the conditional format parameters.
        settings: A dictionary of settings for the report template.
    """

    def __init__(
        self,
        sections: Optional[dict] = None,
        formats: Optional[dict] = None,
        conditional_formats: Optional[dict] = None,
        settings: Optional[dict] = None,
    ):
        """Initialize a ReportTemplate.

        Args:
            sections: A dictionary of sections in the report template.
            formats: A dictionary of formats in the report template.
            conditional_formats: A dictionary of conditional formats in the report
                template.
            settings: A dictionary of settings for the report template.
        """
        document = {
            "sections": {} if sections is None else sections,
            "formats": {} if formats is None else formats,
            "conditional_formats": (
                {} if conditional_formats is None else conditional_formats
            ),
            "settings": {} if settings is None else settings,
        }
        if errors := validate_document_entry_types(document):
            error_message = "\n".join([error.message for error in errors])
            raise ValueError(f"invalid initialization parameters\n{error_message}")

        self.sections = ReportTemplateSections(document["sections"])
        self.formats = ReportTemplateFormats(document["formats"])
        self.conditional_formats = ReportTemplateFormats(
            document["conditional_formats"]
        )
        self.settings = ReportTemplateSettings(document["settings"])

    def to_dict(self) -> dict[str, dict]:
        """Returns a dictionary representation of the `ReportTemplate`."""
        return {
            "sections": self.sections.to_dict(),
            "formats": self.formats.to_dict(),
            "conditional_formats": self.conditional_formats.to_dict(),
            "settings": self.settings.to_dict(),
        }

    @classmethod
    def from_dict(cls, template_document: dict) -> ReportTemplate:
        """Creates a `ReportTemplate` instance from a dictionary.

        Args:
            template_document: A dictionary representation of a `ReportTemplate`. The
                keys "sections", "formats", "conditional_formats", and "settings" are
                used to initialize the `ReportTemplate` instance.

        Returns:
            A `ReportTemplate` instance.
        """
        return cls(
            sections=template_document.get("sections", {}),
            formats=template_document.get("formats", {}),
            conditional_formats=template_document.get("conditional_formats", {}),
            settings=template_document.get("settings", {}),
        )

    @classmethod
    def load(cls, filepath) -> ReportTemplate:
        """Loads a report template YAML file and returns a `ReportTemplate` instance."""
        with open(filepath, "r", encoding="utf-8") as file:
            if errors := validate_template_file_integrity(filepath):
                error_message = "\n".join([error.description for error in errors])
                raise ValueError(f"error loading YAML file\n{error_message}")
            template_data = yaml.safe_load(file)
        return cls.from_dict(template_data)

    def save(self, filepath) -> None:
        """Saves the `ReportTemplate` to a YAML file."""
        with open(filepath, "w", encoding="utf-8") as file:
            yaml.dump(
                self.to_dict(),
                file,
                version=(1, 2),
                sort_keys=False,
                Dumper=IndentDumper,
            )


class IndentDumper(yaml.SafeDumper):
    """Custom YAML dumper to preserve indentation."""

    def increase_indent(self, flow=False, indentless=False):
        return super(IndentDumper, self).increase_indent(flow, False)
