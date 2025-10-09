def to_str(value):
    """
    Convert value to string only if it is already a string and not empty.
    Otherwise, return the value unchanged.
    """
    if isinstance(value, str) and value != "":
        return str(value)
    elif value is None:
        return ""
    else:
        return value

def apostrophe(value):
    """
    Add an apostrophe in front of the value if it starts with = or +.

    Args:
        value (str): The input value to check.

    Returns:
        str: The modified value with an apostrophe if it starts with = or +.
    """
    return f"'{value}" if value.startswith(('=', '+')) else value