"""Module for storing, loading and manipulating table templates.

The `TableTemplate` class is a Python representation of a YAML table template file and
contains the configuration instructions for compiling a table into a formatted Excel.
The `TableTemplate` class provides methods for loading a table template from a YAML
file, and saving the template to a YAML file.
"""

from __future__ import annotations
from typing import Optional

import yaml

from xlsxreport.validate import (
    validate_document_entry_types,
    validate_template_file_integrity,
)
from xlsxreport.template.sections import TableTemplateSections
from xlsxreport.template.settings import TableTemplateSettings
from xlsxreport.template.formats import TableTemplateFormats


class TableTemplate:
    """Representation of a table template and its configuration.

    Attributes:
        sections: A mapping representing the sections of the table template. Each
            key-value pair corresponds to a template section, where the key is the
            section's name and the value is a `TemplateSection` containing the section
            parameters.
        formats: A mapping representing the formats of the table template. Each
            key-value pair corresponds to a format, where the key is the format's name
            and the value is dictionary containing the parameters for that format.
        conditional_formats: A mapping representing the conditional formats of the table
            template. Each key-value pair corresponds to a conditional format, where the
            key is the format's name and the value is dictionary containing the
            parameters for that format.
        settings: A mapping that contains settings for the table template. Each
            key-value pair represents a setting and its value.
    """

    sections: TableTemplateSections
    formats: TableTemplateFormats
    conditional_formats: TableTemplateFormats
    settings: TableTemplateSettings

    def __init__(
        self,
        sections: Optional[dict] = None,
        formats: Optional[dict] = None,
        conditional_formats: Optional[dict] = None,
        settings: Optional[dict] = None,
    ):
        """Initialize a TableTemplate.

        Args:
            sections: A dictionary containing template section descriptions.
            formats: A dictionary containing format descriptions.
            conditional_formats: A dictionary containing conditional format
                descriptions.
            settings: A dictionary containing table template settings.
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

        self.sections = TableTemplateSections(document["sections"])
        self.formats = TableTemplateFormats(document["formats"])
        self.conditional_formats = TableTemplateFormats(document["conditional_formats"])
        self.settings = TableTemplateSettings(document["settings"])

    def to_dict(self) -> dict[str, dict]:
        """Return a dictionary representation of the `TableTemplate`."""
        return {
            "sections": self.sections.to_dict(),
            "formats": self.formats.to_dict(),
            "conditional_formats": self.conditional_formats.to_dict(),
            "settings": self.settings.to_dict(),
        }

    @classmethod
    def from_dict(cls, template_document: dict) -> TableTemplate:
        """Create a `TableTemplate` instance from a dictionary.

        Args:
            template_document: A dictionary representation of a `TableTemplate`. The
                keys "sections", "formats", "conditional_formats", and "settings" are
                used to initialize the `TableTemplate` instance.

        Returns:
            A `TableTemplate` instance.
        """
        return cls(
            sections=template_document.get("sections", {}),
            formats=template_document.get("formats", {}),
            conditional_formats=template_document.get("conditional_formats", {}),
            settings=template_document.get("settings", {}),
        )

    @classmethod
    def load(cls, filepath) -> TableTemplate:
        """Load a table template YAML file and return a `TableTemplate` instance."""
        with open(filepath, "r", encoding="utf-8") as file:
            if errors := validate_template_file_integrity(filepath):
                error_message = "\n".join([error.description for error in errors])
                raise ValueError(f"error loading YAML file\n{error_message}")
            template_data = yaml.safe_load(file)
        return cls.from_dict(template_data)

    def save(self, filepath) -> None:
        """Save the `TableTemplate` to a YAML file."""
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
