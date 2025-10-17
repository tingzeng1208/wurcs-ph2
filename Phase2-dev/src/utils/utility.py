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

def excel_cell_to_col_row(cell_ref):
    """Convert Excel cell reference like 'E8' to (column, row) tuple."""
    col_str = ""
    row_str = ""
    
    for char in cell_ref.strip():
        if char.isalpha():
            col_str += char
        elif char.isdigit():
            row_str += char
    
    # Convert column letters to number (A=1, B=2, etc.)
    col_num = 0
    for char in col_str.upper():
        col_num = col_num * 26 + (ord(char) - ord('A') + 1)
    
    return (col_num, int(row_str))

def scrub_year(s: str, current_year: int) -> str:
        s = s.replace("#Y#", str(current_year))
        s = s.replace("#Y-1#", str(current_year - 1))
        s = s.replace("#Y-2#", str(current_year - 2))
        s = s.replace("#Y-3#", str(current_year - 3))
        s = s.replace("#Y-4#", str(current_year - 4))
        s = s.replace("#Y-5#", str(current_year - 4))
        return s