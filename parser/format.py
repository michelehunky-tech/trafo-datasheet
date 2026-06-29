"""Deterministic formatting of values for the English datasheet."""
from datetime import datetime, date


def is_blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def looks_like_coerced_date(v):
    """Excel sometimes converts slash-formatted strings (e.g. 24/50/125) into dates."""
    return isinstance(v, (datetime, date))


def format_number(v, decimals, thousands=",", decimal_sep="."):
    """Format a numeric value with fixed decimals and thousands separator (English).
    Trailing zeros after the decimal point are stripped (6.00 -> 6, 0.50 -> 0.5)."""
    try:
        num = float(v)
    except (TypeError, ValueError):
        return str(v)
    if decimals is None:
        decimals = 0 if float(num).is_integer() else 2
    s = f"{num:,.{decimals}f}"  # python default: , thousands and . decimal
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    if thousands != "," or decimal_sep != ".":
        s = s.replace(",", "\0").replace(".", decimal_sep).replace("\0", thousands)
    return s


def format_value(raw, field, number_format):
    """Return the display string for a value, applying decimals/thousands/scale when numeric.
    Text values are returned as-is (translation is applied separately via value_map)."""
    if is_blank(raw):
        return None
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        num = raw * field.get("scale", 1)
        return format_number(
            num,
            field.get("decimals"),
            number_format.get("thousands", ","),
            number_format.get("decimal", "."),
        )
    return str(raw).strip()
