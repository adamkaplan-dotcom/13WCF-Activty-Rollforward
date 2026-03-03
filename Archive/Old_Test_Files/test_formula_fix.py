#!/usr/bin/env python3
"""
Test script for improved formula adjustment function
"""

import re

def adjust_formula_columns(formula, column_offset=1):
    """
    Adjust column references in an Excel formula by a given offset.
    Only adjusts relative references, preserves absolute ($) references.
    """
    if not formula or not formula.startswith('='):
        return formula

    def col_to_num(col_letters):
        """Convert column letters to number (A=1, Z=26, AA=27, etc.)"""
        col_num = 0
        for char in col_letters:
            col_num = col_num * 26 + (ord(char) - ord('A') + 1)
        return col_num

    def num_to_col(num):
        """Convert number to column letters (1=A, 26, Z, 27=AA, etc.)"""
        col_letters = ''
        while num > 0:
            num -= 1
            col_letters = chr(ord('A') + (num % 26)) + col_letters
            num //= 26
        return col_letters

    def adjust_cell_ref(match):
        """Adjust a single cell reference"""
        prefix_char = match.group(1) or ''
        dollar1 = match.group(2) or ''
        col_letters = match.group(3)
        dollar2 = match.group(4) or ''
        row_number = match.group(5) or ''

        # Only adjust if column reference is relative (no $ before column)
        if dollar1:
            # Absolute column reference - don't adjust
            return prefix_char + dollar1 + col_letters + dollar2 + row_number
        else:
            # Relative column reference - adjust it
            col_num = col_to_num(col_letters)
            new_col = num_to_col(col_num + column_offset)
            return prefix_char + dollar1 + new_col + dollar2 + row_number

    # Pattern matches cell references with row numbers (not bare column letters)
    # This avoids matching function names like INDEX, MATCH, etc.
    # Matches: A1, $A1, A$1, $A$1, CC14, $CC$14, etc.
    # Also matches ranges like CC14:CC20
    pattern = r'(^|[^A-Za-z])(\$?)([A-Z]+)(\$?)(\d+)'

    # First pass: adjust cell references with row numbers
    adjusted_formula = re.sub(pattern, adjust_cell_ref, formula)

    # Second pass: adjust column-only ranges like $B:$B or A:Z
    def adjust_col_range(match):
        prefix = match.group(1) or ''
        dollar1 = match.group(2) or ''
        col1 = match.group(3)
        colon = match.group(4)
        dollar2 = match.group(5) or ''
        col2 = match.group(6)
        suffix = match.group(7) or ''

        # Adjust first column if relative
        if dollar1:
            new_col1 = col1
        else:
            new_col1 = num_to_col(col_to_num(col1) + column_offset)

        # Adjust second column if relative
        if dollar2:
            new_col2 = col2
        else:
            new_col2 = num_to_col(col_to_num(col2) + column_offset)

        return prefix + dollar1 + new_col1 + colon + dollar2 + new_col2 + suffix

    # Match column ranges like $B:$B, A:Z, etc.
    range_pattern = r'([^A-Za-z])(\$?)([A-Z]+)(:)(\$?)([A-Z]+)([^A-Za-z]|$)'
    adjusted_formula = re.sub(range_pattern, adjust_col_range, adjusted_formula)

    return adjusted_formula


# Test cases
test_cases = [
    ("=CC14+7", "=CD14+7"),
    ("=INDEX($B:$B,MATCH(CC14,$A:$A,0))", "=INDEX($B:$B,MATCH(CD14,$A:$A,0))"),
    ("=$CC$14+7", "=$CC$14+7"),  # Absolute column stays the same when dragging right
    ("=CC$14+CD15", "=CD$14+CE15"),
    ("=SUM(CC14:CC20)", "=SUM(CD14:CD20)"),
    ("=A1+B2+C3", "=B1+C2+D3"),
    ("=VLOOKUP(CC14,$A:$Z,2,FALSE)", "=VLOOKUP(CD14,$A:$Z,2,FALSE)"),
]

print("Testing improved formula adjustment function:\n")
all_passed = True

for original, expected in test_cases:
    result = adjust_formula_columns(original, column_offset=1)
    status = "✓" if result == expected else "✗"

    if result != expected:
        all_passed = False

    print(f"{status} Original:  {original}")
    print(f"  Expected:  {expected}")
    print(f"  Got:       {result}")
    print()

if all_passed:
    print("✓ All tests passed!")
else:
    print("✗ Some tests failed!")
