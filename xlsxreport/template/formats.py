from __future__ import annotations
from collections import UserDict
from copy import deepcopy

from xlsxreport.template._repr import dict_to_string


class ReportTemplateFormats(UserDict):
    """Representation of report template settings."""

    def __init__(self, data: dict[str, dict]):
        all_keys_are_strings = all([isinstance(key, str) for key in data])
        all_values_are_dicts = all([isinstance(value, dict) for value in data.values()])
        if not all_keys_are_strings or not all_values_are_dicts:
            raise TypeError(
                "All format keys must be strings and all values must be dictionaries."
            )
        self.data: dict = data

    def __repr__(self):
        return _format_formats(self.data, double_quotes=True)

    def to_dict(self) -> dict:
        """Return a copy of the formats as a dictionary."""
        return deepcopy(self.data)


def _format_formats(
    formats: dict, indent: int = 4, line_length: int = 80, double_quotes: bool = False
) -> str:
    item_strings = []
    for key, value in formats.items():
        prefix = f"{key}: "
        lines = dict_to_string(value, indent, line_length, double_quotes, prefix)
        item_strings.append("\n".join(lines))
    return "\n".join(item_strings)
