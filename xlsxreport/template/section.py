from __future__ import annotations
from copy import deepcopy
from enum import Enum
from typing import Any

import cerberus  # type: ignore

from xlsxreport.template._repr import dict_to_string


OPTIONAL_SECTION_PARAMS: dict[str, dict[str, str | float | bool]] = {
    "format": {"type": "string"},
    "column_format": {"type": "dict"},
    "conditional_format": {"type": "string"},
    "column_conditional_format": {"type": "dict"},
    "header_format": {"type": "dict"},
    "supheader": {"type": "string"},
    "supheader_format": {"type": "dict"},
    "width": {"type": "float"},
    "border": {"type": "boolean", "default": False},
}


STANDARD_SECTION_SCHEMA: dict[str, dict[str, str | float | bool]] = {
    "columns": {"required": True, "type": "list"},
    **OPTIONAL_SECTION_PARAMS,
}


TAG_SECTION_SCHEMA: dict[str, dict[str, str | float | bool]] = {
    "tag": {"required": True, "type": "string"},
    "remove_tag": {"type": "boolean", "default": False},
    "log2": {"type": "boolean", "default": False},
    **OPTIONAL_SECTION_PARAMS,
}

LABEL_TAG_SECTION_SCHEMA: dict[str, dict[str, str | float | bool]] = {
    "tag": {"required": True, "type": "string"},
    "labels": {"required": True, "type": "list"},
    "remove_tag": {"type": "boolean", "default": False},
    "log2": {"type": "boolean", "default": False},
    **OPTIONAL_SECTION_PARAMS,
}


COMPARISON_SECTION_SCHEMA: dict[str, dict[str, str | float | bool]] = {
    "comparison_group": {"required": True, "type": "boolean"},
    "tag": {"required": True, "type": "string"},
    "columns": {"required": True, "type": "list"},
    "replace_comparison_tag": {"type": "string"},
    "remove_tag": {"type": "boolean", "default": False},
    **OPTIONAL_SECTION_PARAMS,
}


class SectionCategory(Enum):
    """Enum for section categories."""

    UNKNOWN = -1
    STANDARD = 1
    TAG = 2
    LABEL_TAG = 3
    COMPARISON = 4


_template_section_schemas = {
    SectionCategory.UNKNOWN: OPTIONAL_SECTION_PARAMS,
    SectionCategory.STANDARD: STANDARD_SECTION_SCHEMA,
    SectionCategory.TAG: TAG_SECTION_SCHEMA,
    SectionCategory.LABEL_TAG: LABEL_TAG_SECTION_SCHEMA,
    SectionCategory.COMPARISON: COMPARISON_SECTION_SCHEMA,
}


class TemplateSection:
    """Representation of a table section."""

    def __init__(self, data: dict):
        if not isinstance(data, dict):
            raise TypeError("Section data must be a dictionary")
        self.category = _identify_section_category(data)
        self.schema = deepcopy(_template_section_schemas[self.category])
        self.data = data

        self._validator = cerberus.Validator(require_all=False, allow_unknown=False)

    def __contains__(self, key: str) -> bool:
        return key in self.data

    def __getitem__(self, key: str) -> str | float | bool | list | dict:
        if key not in self.schema:
            raise KeyError(f"Invalid {self.category.name} section parameter '{key}'")
        if key not in self.data:
            raise KeyError(f"Section parameter '{key}' not defined")
        return self.data[key]

    def __setitem__(self, key: str, value: str | float | bool | list | dict) -> None:
        updated_data = self.to_dict()
        updated_data.update({key: value})
        if not self._validator.validate(updated_data, self.schema):
            raise ValueError(
                f"Invalid {self.category.name} section parameter: "
                f"{self._validator.errors}"
            )

        self.data[key] = value

    def __repr__(self) -> str:
        prefix = f"{self.category.name} section: "
        lines = dict_to_string(
            self.data,
            indent=4,
            line_length=80,
            double_quotes=True,
            prefix=prefix,
        )
        return "\n".join(lines)

    def get(self, key: str, default: Any = None) -> Any:
        """Get a section parameter or return a default value if not found."""
        try:
            return self.__getitem__(key)
        except KeyError:
            return default

    def to_dict(self) -> dict:
        """Return a copy of the section as a dictionary."""
        return deepcopy(self.data)


def _identify_section_category(section: dict) -> SectionCategory:
    """Use section schemas to identify the category of a section."""
    validator = cerberus.Validator()
    validator.allow_unknown = False
    validator.require_all = False
    matched_categories = []
    for category, schema in _template_section_schemas.items():
        if validator.validate(section, schema):
            matched_categories.append(category)

    if len(matched_categories) > 1:
        raise ValueError(
            f"Section matched to multiple categories: {matched_categories}"
        )

    if not matched_categories:
        return SectionCategory.UNKNOWN
    return matched_categories[0]
