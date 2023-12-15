"""Command to validate a YAML template file."""
import click
import yaml

from xlsxreport import get_template_path
from xlsxreport.validate import (
    ErrorLevel,
    validate_template_file_integrity,
    validate_document_entry_types,
    validate_template_content,
)


@click.command()
@click.argument("template")
def validate_command(template: str):
    """Validate a YAML template file and print detected errors to the console.

    The TEMPLATE argument is first used to look for a file with the specified filepath.
    If no file is found, the XlsxReport appdata directory is searched for a file with
    the corresponding name.
    """
    template_path = get_template_path(template)
    if template_path is None:
        raise click.ClickException(
            f"Invalid value for 'TEMPLATE': Path '{click.format_filename(template)}' "
            "does not exist."
        )

    click.echo(f"Validating YAML template file: {click.format_filename(template_path)}")

    if integrity_errors := validate_template_file_integrity(template_path):
        output = ["Error loading YAML file, validation cannot proceed."]
        output.extend([err.message for err in integrity_errors])
        click.echo("\n".join(output))
        return

    with open(template_path, "r", encoding="utf-8") as file:
        template_document = yaml.safe_load(file)

    if type_errors := validate_document_entry_types(template_document):
        output = ["Type errors detected, validation cannot proceed."]
        output.extend([err.message for err in type_errors])
        click.echo("\n".join(output))
        return

    if content_errors := validate_template_content(template_document):
        max_error_level = max((err.error_level for err in content_errors))
        if max_error_level <= ErrorLevel.INFO:
            output = ["Template is valid for Excel report generation."]
        elif max_error_level <= ErrorLevel.WARNING:
            output = [
                "Only non-serious issues detected, template is valid for Excel report "
                "generation."
            ]
        elif max_error_level <= ErrorLevel.ERROR:
            output = [
                "Errors detected, template is usable for Excel report generation but "
                "might lead to an unexpected result."
            ]
        else:
            raise ValueError("Template contains unexpected critical errors.")
        output.extend([err.message for err in content_errors])
        click.echo("\n".join(output))
        return

    click.echo("Template is valid for Excel report generation.")
