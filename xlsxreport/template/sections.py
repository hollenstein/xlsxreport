from __future__ import annotations
from collections import UserDict
from typing import Any

from xlsxreport.template.section import TemplateSection


class TableTemplateSections(UserDict):
    """Container for table template section descriptions."""

    __marker = object()

    def __init__(self, sections: dict[str, dict]):
        self.data: dict[str, TemplateSection] = {
            k: TemplateSection(v) for k, v in sections.items()
        }

    def __getitem__(self, key: str | int) -> TemplateSection:
        if isinstance(key, str):
            return self.data[key]
        if isinstance(key, int):
            return self.data[_get_section_name_from_position(self.data, key)]

    def __setitem__(self, key: str, value: dict | TemplateSection) -> None:
        self.add(key, value)

    def __repr__(self):
        length = max([len(key) for key in self.data])
        section_count = len(self.data)
        num_digits = len(str(section_count - 1))

        section_strings = []
        for i, (name, section) in enumerate(self.data.items(), 0):
            trailing_name_spaces = " " * (length - len(name))
            section_strings.append(
                f"[{i:{num_digits}}] "
                f"{name}:{trailing_name_spaces} "
                f"{section.category.name} section"
            )
        return "\n".join(section_strings)

    def reposition(self, section: str | int, to: str | int) -> None:
        """Move the indicated section to the specified position.

        Args:
            section: The name or index of the section to move.
            to: The index to move the `section` to or the name of a section to move the
                the `section` to. If a negative position is specified it is counted from
                the end of the dict. I.e. -1 will insert the key-value pair before the
                last entry in the dict. If a secttion name is specified, the `section`
                is moved before the specified section.
        """
        if isinstance(section, int):
            section_name = _get_section_name_from_position(self.data, section)
        else:
            section_name = section

        if isinstance(to, int):
            self.data = _move_key_to_position(self.data, section_name, to)
        else:
            self.data = _switch_key_positions(self.data, section_name, to)

    def add(self, section_name: str, section: dict | TemplateSection) -> None:
        """Add a new section to the template or overwrite an existing section.

        Args:
            section_name: The name of the new section.
            section: The parameters of the new section or a TemplateSection object.
        """
        if not isinstance(section, TemplateSection):
            self.data[section_name] = TemplateSection(section)
        else:
            self.data[section_name] = section

    def insert(
        self, section_name: str, section: dict | TemplateSection, to: str | int
    ) -> None:
        """Insert a new section into the template at the specified position.

        Args:
            section_name: The name of the new section.
            section: The parameters of the new section or a TemplateSection object.
            to: The index to insert the new section at or the name of the section to
                insert the new section before. If a negative position is specified it is
                counted from the end of the dict. I.e. -1 will insert the key-value pair
                before the last entry in the dict. If a secttion name is specified, the
                new section is inserted before the specified section.
        """
        if not isinstance(section, TemplateSection):
            template_section = TemplateSection(section)
        else:
            template_section = section

        if isinstance(to, int):
            self.data = _insert_item_at_position(
                self.data, section_name, template_section, to
            )
        else:
            self.data = _insert_item_before_key(
                self.data, section_name, template_section, to
            )

    def pop(self, section: str | int, default: Any = __marker) -> TemplateSection | Any:
        """Remove the indicated section from the template and return it.

        Args:
            section: The name or index of the section to remove.
            default: The value to return if the section is not found.

        Returns:
            The removed section.
        """
        if isinstance(section, int):
            try:
                section_name = _get_section_name_from_position(self.data, section)
            except ValueError:
                if default is self.__marker:
                    raise
                return default
        else:
            if section not in self.data:
                if default is self.__marker:
                    raise KeyError(section)
                return default
            section_name = section

        return self.data.pop(section_name)

    def remove(self, section: str | int) -> None:
        """Remove the indicated section from the template.

        Args:
            section: The name or index of the section to remove.
        """
        if isinstance(section, int):
            section_name = _get_section_name_from_position(self.data, section)
        else:
            section_name = section

        del self.data[section_name]

    def to_dict(self) -> dict[str, dict]:
        """Return a copy of the sections as dictionaries."""
        return {k: v.to_dict() for k, v in self.data.items()}


def _get_section_position_from_name(sections: dict[str, Any], key: str) -> int:
    if key not in sections:
        raise ValueError(f"Invalid section name '{key}'")
    return list(sections).index(key)


def _get_section_name_from_position(sections: dict[str, Any], index: int) -> str:
    if index < len(sections) * -1 or index > len(sections) - 1:
        raise ValueError(f"Index {index} out of range")
    return list(sections)[index]


def _switch_key_positions(
    _dict: dict[str, Any], key1: str, key2: str
) -> dict[str, Any]:
    """Move key1 in front of key2 in the dictionary."""
    if key1 not in _dict:
        raise ValueError(f"Invalid key '{key1}'")
    if key2 not in _dict:
        raise ValueError(f"Invalid key '{key2}'")
    if key1 == key2:
        return _dict
    key_order = [k for k in _dict if k != key1]
    index = list(key_order).index(key2)
    key_order.insert(index, key1)
    return {k: _dict[k] for k in key_order}


def _move_key_to_position(
    _dict: dict[str, Any], key: str, position: int
) -> dict[str, Any]:
    """Move the indicated key to the specified position in the dictionary."""
    if key not in _dict:
        raise ValueError(f"Invalid key '{key}'")
    if position < len(_dict) * -1 or position > len(_dict) - 1:
        raise ValueError(f"Index {position} out of range")

    key_order = [k for k in _dict if k != key]
    index = position if position >= 0 else len(_dict) + position
    key_order.insert(index, key)
    return {k: _dict[k] for k in key_order}


def _insert_item_at_position(
    _dict: dict[str, Any], key: str, value: Any, index: int
) -> dict[str, Any]:
    """Insert the key-value pair before the specified index in the dictionary.

    If the index is negative, it is counted from the end of the dict. I.e. -1 will
    insert the key-value pair before the last entry in the dict.
    """
    if index < len(_dict) * -1 or index > len(_dict):
        raise ValueError("Index out of range")

    index = index if index >= 0 else len(_dict) + index
    keys = list(_dict)
    keys.insert(index, key)
    values = list(_dict.values())
    values.insert(index, value)
    return dict(zip(keys, values))


def _insert_item_before_key(
    _dict: dict[str, Any], key: str, value: Any, old_key: str
) -> dict[str, Any]:
    """Insert the new key-value pair before the specified key in the dictionary."""
    if old_key not in _dict:
        raise ValueError(f"Invalid key '{old_key}'")
    keys = list(_dict)
    index = keys.index(old_key)
    keys.insert(index, key)
    values = list(_dict.values())
    values.insert(index, value)
    return dict(zip(keys, values))
