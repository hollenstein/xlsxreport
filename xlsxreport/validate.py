"""Module for validation of table templates and their YAML file representation.

Main API functions:
- validate_template_file_integrity(filepath: str)
- validate_document_entry_types(template_document: dict)
- validate_template_content(template_document: dict)
"""

from __future__ import annotations
from dataclasses import dataclass
from enum import Enum, IntEnum
import logging
from typing import Any
import warnings

import cerberus  # type: ignore
import xlsxwriter  # type: ignore
import yaml

from xlsxreport.schemas import TEMPLATE_SCHEMA, SETTINGS_SCHEMA


class MainSections(Enum):
    SECTIONS = "sections"
    FORMATS = "formats"
    CONDITIONAL_FORMATS = "conditional_formats"
    SETTINGS = "settings"


SPECIAL_FORMATS = ["header", "supheader"]


class ErrorLevel(IntEnum):
    INFO = logging.INFO
    WARNING = logging.WARNING
    ERROR = logging.ERROR
    CRITICAL = logging.CRITICAL


class ValidationErrorType(Enum):
    UNKNOWN_PARAMTER = "unknown parameter"
    MISSING_PARAMETER = "missing parameter"
    INVALID_VALUE_TYPE = "invalid value type"
    INVALID_FORMAT = "invalid format description"
    UNDEFINED_FORMAT = "format not defined"
    UNUSED_FORMAT = "format not used"
    TYPE_ERROR = "type error"
    FILE_ERROR = "file error"


@dataclass
class ValidationError:
    error_type: ValidationErrorType
    error_level: ErrorLevel
    field: tuple[str, ...]
    description: str

    def __post_init__(self):
        if isinstance(self.field, str):
            raise Exception()

    @property
    def message(self) -> str:
        field_string = ".".join([str(f) for f in self.field])
        return f"[{self.error_level.name}] {field_string}: {self.description}"

    def __repr__(self) -> str:
        return self.message


# MAIN API FUNCTIONS
def validate_template_file_integrity(filepath: str) -> list[ValidationError]:
    """Validate the integrity of a table template YAML file.

    All reported errors are of level CRITICAL, since a YAML template containing file
    integrity errors cannot be imported as a `xlsxreport.TableTemplate`.

    This function checks if the specified filepath can be loaded as a YAML file and if
    the root of the YAML document is a dictionary. It does currently not check for
    duplicated keys in the YAML document.

    Args:
        filepath: The path to a YAML file.

    Returns:
        A list of `ValidationError`s of type FILE_ERROR and level CRITICAL. Each entry
        corresponds to a problem with the specified filepath. If the list is empty, the
        filepath can be loaded as a YAML file and the root of the YAML document is a
        dictionary.
    """
    if errors := validate_template_file_loading(filepath):
        return errors

    template_document = yaml.safe_load(open(filepath, "r", encoding="utf-8"))
    if errors := validate_template_document_root_type(template_document):
        return errors

    return []


def validate_document_entry_types(template_document: dict) -> list[ValidationError]:
    """Check if the types of the values in the YAML document are correct.

    All reported type errors are of level CRITICAL, since a table template containing
    type errors will result in the Excel table generation to fail.

    Args:
        template_document: A dictionary representation of a `TableTemplate`. The values
            of the keys "sections", "formats", "conditional_formats", and "settings" are
            validated according to the `TEMPLATE_SCHEMA`.

    Returns:
        A list of `ValidationError`s of type TYPE_ERROR and level CRITICAL. Each entry
        corresponds to a value in the YAML document that has an incorrect type. If the
        list is empty, the types of the values in the YAML document are correct.
    """
    validator = cerberus.Validator()
    validator.allow_unknown = True
    validator.require_all = False
    validator.validate(template_document, TEMPLATE_SCHEMA)

    errors = []
    for field, description in _flatten_cerberus_errors(validator.errors):
        error = ValidationError(
            ValidationErrorType.TYPE_ERROR, ErrorLevel.CRITICAL, field, description
        )
        errors.append(error)
    return errors


def validate_template_content(template_document: dict) -> list[ValidationError]:
    """Validate the content of a table template YAML file.

    Reported `ValidationError`s are of level INFO, WARNING or ERROR, depending on the
    severity of the problem. INFO level errors do not indicate any problems but are only
    reported for information. WARNING level errors indicate that the template content is
    not optimal. ERROR level errors indicate that the template content is not valid. A
    table template containing ERROR level errors can still be imported as a
    `TableTemplate` and be used to generate an Excel table, however, the Excel table
    generation will not work as expected.

    Args:
        template_document: A dictionary representation of a `TableTemplate`. The keys
            "sections", "formats", "conditional_formats", and "settings" are used to
            validate the template content.

    Returns:
        A list of `ValidationError`s. Each entry corresponds to a problem with the
        template content. If the list is empty, no problems were found.
    """
    formats_section = template_document.get(MainSections.FORMATS.value, {})
    conditional_formats_section = template_document.get(
        MainSections.CONDITIONAL_FORMATS.value, {}
    )
    settings_section = template_document.get(MainSections.SETTINGS.value, {})

    error_lists = [
        # main sections
        validate_expected_main_sections(template_document),
        validate_unexpected_main_sections(template_document),
        # individual sections
        validate_format_descriptions(formats_section),
        validate_conditional_format_descriptions(conditional_formats_section),
        validate_expected_settings_parameters(settings_section),
        validate_unexpected_settings_parameters(settings_section),
        # Todo - validation of section entries is missing
        # 1) check if all sections contain only known parameters (WARNING)
        # 2) check if all sections can be assigned a known category (ERROR)
        # Could also get a "formats" and a "sections" dictionary instead of a template_document
        validate_unused_formats(template_document),
        validate_undefined_formats(template_document),
        validate_special_formats_defined(template_document),
        validate_unused_conditional_formats(template_document),
        validate_undefined_conditional_formats(template_document),
    ]
    errors = [item for error_list in error_lists for item in error_list]
    return errors


# YAML INTEGRITY VALIDATION
def validate_template_file_loading(filepath: str) -> list[ValidationError]:
    """Check if the specified filepath can be loaded as a YAML file

    Args:
        filepath: The path to a YAML file.

    Returns:
        A list of `ValidationError`s of type FILE_ERROR and level CRITICAL. Each entry
        corresponds to a problem with the specified file. If the list is empty, the
        filepath can be loaded as a YAML file.
    """
    try:
        with open(filepath, "r", encoding="utf-8") as file:
            _ = yaml.safe_load(file)
    except yaml.scanner.ScannerError:
        description = f"invalid syntax, cannot parse file '{filepath}'"
        error = ValidationError(
            ValidationErrorType.FILE_ERROR,
            ErrorLevel.CRITICAL,
            ("yaml file",),
            description,
        )
        return [error]
    except yaml.error.MarkedYAMLError:
        raise NotImplementedError
    except yaml.error.YAMLError:
        raise NotImplementedError

    return []


def validate_template_document_root_type(
    template_document: dict | Any,
) -> list[ValidationError]:
    """Check if the root of the YAML document is a dictionary.

    Args:
        template_document: An imported YAML document.

    Returns:
        If the template document is not a dictionary, a list containing a single
        `ValidationError` is returned. The error type is TYPE_ERROR and the error level
        is CRITICAL. An empty list is returned if the template document is a dictionary.
    """
    if not isinstance(template_document, dict):
        field = ("root",)
        description = (
            f"The YAML document must be a dictionary at its root, not a "
            f"'{type(template_document).__name__}'"
        )
        error = ValidationError(
            ValidationErrorType.TYPE_ERROR, ErrorLevel.CRITICAL, field, description
        )
        return [error]
    return []


# VALIDATE WHICH MAIN SECTIONS ARE PRESENT
def validate_expected_main_sections(template_document: dict) -> list[ValidationError]:
    """Check if all expected main sections are present."""
    errors = []
    for main_section in [e.value for e in MainSections]:
        if main_section not in template_document:
            error = ValidationError(
                ValidationErrorType.MISSING_PARAMETER,
                ErrorLevel.WARNING,
                (main_section,),
                "Expected main section missing",
            )
            errors.append(error)
    return errors


def validate_unexpected_main_sections(template_document: dict) -> list[ValidationError]:
    """Check if only allowed main sections are present."""
    errors = []
    for main_section in template_document:
        if main_section not in [e.value for e in MainSections]:
            error = ValidationError(
                ValidationErrorType.UNKNOWN_PARAMTER,
                ErrorLevel.WARNING,
                (main_section,),
                "Unexpected main section",
            )
            errors.append(error)
    return errors


# FORMAT VALIDATION
def validate_unused_formats(template_document: dict) -> list[ValidationError]:
    """Check if all defined formats, except special ones, are used in the template."""

    used_formats = _retrieve_used_formats(
        template_document.get(MainSections.SECTIONS.value, {})
    )
    defined_formats = set(template_document.get(MainSections.FORMATS.value, {}))
    unused_formats = defined_formats.difference(used_formats)
    unused_formats = unused_formats.difference(SPECIAL_FORMATS)

    errors = []
    for unused_format in unused_formats:
        error = ValidationError(
            ValidationErrorType.UNUSED_FORMAT,
            ErrorLevel.INFO,
            (MainSections.FORMATS.value, unused_format),
            "Format defined but not used",
        )
        errors.append(error)
    return errors


def validate_undefined_formats(template_document: dict) -> list[ValidationError]:
    """Check if all used formats are defined in the template."""
    used_formats = _retrieve_used_formats(
        template_document.get(MainSections.SECTIONS.value, {})
    )
    defined_formats = set(template_document.get(MainSections.FORMATS.value, {}))
    undefined_formats = used_formats.difference(defined_formats)

    errors = []
    for undefined_format in undefined_formats:
        error = ValidationError(
            ValidationErrorType.UNDEFINED_FORMAT,
            ErrorLevel.ERROR,
            (MainSections.FORMATS.value, undefined_format),
            "Format referenced but not defined, falling back to the default format",
        )
        errors.append(error)
    return errors


def validate_special_formats_defined(template_document: dict) -> list[ValidationError]:
    """Check if the special formats are defined in the template."""
    defined_formats = set(template_document.get(MainSections.FORMATS.value, {}))
    undefined_special_formats = set(SPECIAL_FORMATS).difference(defined_formats)

    errors = []
    for undefined_format in undefined_special_formats:
        error = ValidationError(
            ValidationErrorType.UNDEFINED_FORMAT,
            ErrorLevel.INFO,
            (MainSections.FORMATS.value, undefined_format),
            "Format not defined, falling back to the default format",
        )
        errors.append(error)
    return errors


def validate_unused_conditional_formats(
    template_document: dict,
) -> list[ValidationError]:
    """Check if all defined conditional formats are used in the template."""
    used_formats = _retrieve_used_conditional_formats(
        template_document.get(MainSections.SECTIONS.value, {})
    )
    defined_formats = set(
        template_document.get(MainSections.CONDITIONAL_FORMATS.value, {})
    )
    unused_formats = defined_formats.difference(used_formats)

    errors = []
    for unused_format in unused_formats:
        error = ValidationError(
            ValidationErrorType.UNUSED_FORMAT,
            ErrorLevel.INFO,
            (MainSections.CONDITIONAL_FORMATS.value, unused_format),
            "Conditional format defined but not used",
        )
        errors.append(error)
    return errors


def validate_undefined_conditional_formats(
    template_document: dict,
) -> list[ValidationError]:
    """Check if all used conditional formats are defined in the template."""
    used_formats = _retrieve_used_conditional_formats(
        template_document.get(MainSections.SECTIONS.value, {})
    )
    defined_formats = set(
        template_document.get(MainSections.CONDITIONAL_FORMATS.value, {})
    )
    undefined_formats = used_formats.difference(defined_formats)

    errors = []
    for undefined_format in undefined_formats:
        error = ValidationError(
            ValidationErrorType.UNDEFINED_FORMAT,
            ErrorLevel.ERROR,
            (MainSections.CONDITIONAL_FORMATS.value, undefined_format),
            "Conditional format referenced but not defined, format will not be applied",
        )
        errors.append(error)
    return errors


def validate_format_descriptions(formats_section: dict) -> list[ValidationError]:
    """Check if all format descriptions match valid Excel formats."""
    workbook = xlsxwriter.Workbook()
    errors = []
    for format_name, format_description in formats_section.items():
        try:
            workbook.add_format(format_description)
        except (AttributeError, ValueError):
            error = ValidationError(
                ValidationErrorType.INVALID_FORMAT,
                ErrorLevel.ERROR,
                (MainSections.FORMATS.value, format_name),
                f"Invalid format description",
            )
            errors.append(error)
    return errors


def validate_conditional_format_descriptions(
    formats_section: dict,
) -> list[ValidationError]:
    """Check if all conditional format descriptions match valid Excel formats."""
    workbook = xlsxwriter.Workbook()
    worksheet = workbook.add_worksheet()
    errors = []
    for format_name, format_description in formats_section.items():
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("error")
                worksheet.conditional_format("A1:B2", format_description)
        except (AttributeError, Warning):
            error = ValidationError(
                ValidationErrorType.INVALID_FORMAT,
                ErrorLevel.ERROR,
                (MainSections.CONDITIONAL_FORMATS.value, format_name),
                f"Invalid format description",
            )
            errors.append(error)
    return errors


# SETTINGS VALIDATION
def validate_expected_settings_parameters(
    settings_section: dict,
) -> list[ValidationError]:
    """Check if only allowed settings parameters are present."""
    errors = []
    for field_name, description in SETTINGS_SCHEMA.items():
        if field_name not in settings_section:
            message = (
                "Missing parameter, falling back to the default value "
                f"`{description['default']}`"
            )
            error = ValidationError(
                ValidationErrorType.MISSING_PARAMETER,
                ErrorLevel.INFO,
                (MainSections.SETTINGS.value, field_name),
                message,
            )
            errors.append(error)
    return errors


def validate_unexpected_settings_parameters(
    settings_section: dict,
) -> list[ValidationError]:
    """Check if all settings parameters are allowed."""
    errors = []
    for field_name in settings_section:
        if field_name not in SETTINGS_SCHEMA:
            error = ValidationError(
                ValidationErrorType.UNKNOWN_PARAMTER,
                ErrorLevel.WARNING,
                (MainSections.SETTINGS.value, field_name),
                f"Unknown parameter",
            )
            errors.append(error)
    return errors


def _flatten_cerberus_errors(errors: dict, field: tuple = ()):
    """Flatten the error dictionary returned by cerberus into a list of tuples.

    Args:
        errors: The error dictionary returned by cerberus.
        field: The field that is currently being validated.

    Returns:
        A list of tuples. Each tuple contains the field that is being validated and the
        error message for that field. A field is a tuple of strings, each string
        corresponding to a level in the dictionary hierarchy to the entry that is being
        validated.
    """
    error_fields = []
    for key, value in errors.items():
        if isinstance(value, list):
            for item in value:
                if isinstance(item, dict):
                    error_fields.extend(_flatten_cerberus_errors(item, field + (key,)))
                else:
                    error_fields.append((field + (key,), item))
        elif isinstance(value, dict):
            error_fields.extend(_flatten_cerberus_errors(value, field + (key,)))
    return error_fields


def _retrieve_used_formats(sections_section: dict) -> set[str]:
    used_formats = set()
    for section in sections_section.values():
        if "format" in section:
            used_formats.add(section["format"])
        if "column_formats" in section:
            used_formats.update(section["column_formats"].values())
    return used_formats


def _retrieve_used_conditional_formats(sections_section: dict) -> set[str]:
    used_formats = set()
    for section in sections_section.values():
        if "conditional_format" in section:
            used_formats.add(section["conditional_format"])
        if "column_conditional_format" in section:
            used_formats.update(section["column_conditional_format"].values())
    return used_formats
