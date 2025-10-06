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