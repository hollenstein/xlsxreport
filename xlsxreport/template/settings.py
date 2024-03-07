from __future__ import annotations
from collections import UserDict
from copy import deepcopy

import cerberus  # type: ignore

from xlsxreport.schemas import SETTINGS_SCHEMA


class ReportTemplateSettings(UserDict):
    """Representation of report template settings."""

    def __init__(self, data: dict):
        self._validator = cerberus.Validator(require_all=False, allow_unknown=False)
        self._schema = SETTINGS_SCHEMA

        settings = {key: value for key, value in data.items() if key in self._schema}
        if not self._validator.validate(settings, self._schema):
            raise TypeError(f"Invalid settings: {self._validator.errors}")

        self.data: dict = settings

    def __getitem__(self, key: str) -> str | float | bool:
        if key not in self._schema:
            raise KeyError(f"Invalid setting argument: {key}")

        if key in self.data:
            return self.data[key]
        else:
            return self._schema[key]["default"]

    def __repr__(self):
        length = max([len(key) for key in self._schema])

        output = []
        for parameter in self._schema:
            value = _format_value(self.get(parameter), quote_char='"')
            if parameter not in self.data:
                value = f"{value} (default)"
            output.append(f"{parameter:<{length}} : {value}")
        return "\n".join(output)

    @property
    def schema(self):
        return deepcopy(self._schema)

    def to_dict(self) -> dict:
        """Return a copy of the settings as a dictionary."""
        return deepcopy(self.data)


def _format_value(value, quote_char):
    if isinstance(value, str):
        return f"{quote_char}{value}{quote_char}"
    else:
        return str(value)
