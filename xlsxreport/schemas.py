"""Schemas defining the layout and structure of the YAML template file.

- `TEMPLATE_SCHEMA`: Defines which main sections are allowed in the template and which
  types of entries are allowed within each section
- `SETTINGS_SCHEMA`: Defines the allowed parameters in the settings section, including
  parameter types and their default values.
- `SECTION_SCHEMA`: Defines the allowed parameters for individual entries within the
  "sections" main section, including the parameter types and for some a default value.
"""

from __future__ import annotations
from typing import Any


SETTINGS_SCHEMA: dict[str, dict[str, str | float | bool]] = {
    "supheader_height": {"type": "float", "min": 0, "default": 20},
    "header_height": {"type": "float", "min": 0, "default": 20},
    "column_width": {"type": "float", "min": 0, "default": 64},
    "log2_tag": {"type": "string", "default": ""},
    "append_remaining_columns": {"type": "boolean", "default": False},
    "write_supheader": {"type": "boolean", "default": False},
    "evaluate_log2_transformation": {"type": "boolean", "default": False},
    "remove_duplicate_columns": {"type": "boolean", "default": True},
    "add_autofilter": {"type": "boolean", "default": True},
    "freeze_cols": {"type": "integer", "min": 0, "default": 1},
}


SECTION_SCHEMA: dict[str, dict[str, str | float | bool]] = {
    "format": {"type": "string"},
    "column_format": {"type": "dict"},
    "conditional_format": {"type": "string"},
    "column_conditional_format": {"type": "dict"},
    "header_format": {"type": "dict"},
    "supheader": {"type": "string"},
    "supheader_format": {"type": "dict"},
    "width": {"type": "float"},
    "border": {"type": "boolean", "default": False},
    "columns": {"type": "list"},
    "tag": {"type": "string"},
    "labels": {"type": "list"},
    "remove_tag": {"type": "boolean", "default": False},
    "log2": {"type": "boolean", "default": False},
    "comparison_group": {"type": "boolean", "default": False},
    "replace_comparison_tag": {"type": "string"},
}


TEMPLATE_SCHEMA: dict[str, dict[str, Any]] = {
    "sections": {
        "type": "dict",
        "keysrules": {"type": "string"},
        "valuesrules": {"type": "dict", "schema": SECTION_SCHEMA},
    },
    "formats": {
        "type": "dict",
        "keysrules": {"type": "string"},
        "valuesrules": {"type": "dict"},
    },
    "conditional_formats": {
        "type": "dict",
        "keysrules": {"type": "string"},
        "valuesrules": {"type": "dict"},
    },
    "settings": {
        "type": "dict",
        "keysrules": {"type": "string"},
        "schema": SETTINGS_SCHEMA,
    },
}
