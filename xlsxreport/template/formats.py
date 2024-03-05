from __future__ import annotations
from collections import UserDict
from copy import deepcopy
import functools


class TemplateFormats(UserDict):
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
        lines = _dict_to_string(value, indent, line_length, double_quotes, prefix)
        item_strings.append("\n".join(lines))
    return "\n".join(item_strings)


def _dict_to_string(
    _dict: dict,
    indent: int,
    line_length: int,
    double_quotes: bool,
    prefix: str = "",
) -> list[str]:
    quote_char = '"' if double_quotes else "'"

    single_line = _single_line_format(_dict, quote_char, prefix)
    if len(single_line) <= line_length:
        return [single_line]

    return _multi_line_format(_dict, quote_char, prefix, indent)


def _single_line_format(_dict: dict, quote_char: str, prefix: str) -> str:
    _format = functools.partial(_format_value, quote_char=quote_char)
    items_string = ", ".join([f"{_format(k)}: {_format(v)}" for k, v in _dict.items()])
    string = f"{prefix}{{{items_string}}}"
    return string


def _multi_line_format(
    _dict: dict, quote_char: str, prefix: str, indent: int
) -> list[str]:
    _format = functools.partial(_format_value, quote_char=quote_char)
    items = [f"{indent * ' '}{_format(k)}: {_format(v)}" for k, v in _dict.items()]
    items = [f"{item}," for item in items[:-1]] + [items[-1]]
    return [f"{prefix}{{", *items, "}"]


def _format_value(value, quote_char):
    if isinstance(value, str):
        return f"{quote_char}{value}{quote_char}"
    else:
        return str(value)
