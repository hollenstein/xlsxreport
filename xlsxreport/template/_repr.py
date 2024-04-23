import functools


def dict_to_string(
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
