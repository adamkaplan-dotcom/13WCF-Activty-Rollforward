#!/usr/bin/env python3
"""
Test script for formula adjustment function
"""

import re

def adjust_formula_columns(formula, column_offset=1):
    """
    Adjust column references in an Excel formula by a given offset.
    """
    if not formula or not formula.startswith('='):
        return formula

    def adjust_cell_ref(match):
        """Adjust a single cell reference"""
        prefix_char = match.group(1) or ''
        dollar1 = match.group(2) or ''
        col_letters = match.group(3)
        dollar2 = match.group(4) or ''
        row_or_range = match.group(5) or ''

        # Convert column letters to number
        col_num = 0
        for char in col_letters:
            col_num = col_num * 26 + (ord(char) - ord('A') + 1)

        # Adjust by offset
        new_col_num = col_num + column_offset

        # Convert back to letters
        new_col_letters = ''
        while new_col_num > 0:
            new_col_num -= 1
            new_col_letters = chr(ord('A') + (new_col_num % 26)) + new_col_letters
            new_col_num //= 26

        return prefix_char + dollar1 + new_col_letters + dollar2 + row_or_range

    pattern = r'(^|[^A-Za-z])(\$?)([A-Z]+)(\$?)(\d+|:\$?[A-Z]+\$?\d*)'
    adjusted_formula = re.sub(pattern, adjust_cell_ref, formula)
    return adjusted_formula


# Test cases
test_cases = [
    ("=CC14+7", "=CD14+7"),
    ("=INDEX($B:$B,MATCH(CC14,$A:$A,0))", "=INDEX($B:$B,MATCH(CD14,$A:$A,0))"),
    ("=$CC$14+7", "=$CD$14+7"),
    ("=CC$14+CD15", "=CD$14+CE15"),
    ("=SUM(CC14:CC20)", "=SUM(CD14:CD20)"),
    ("=A1+B2+C3", "=B1+C2+D3"),
    ("=VLOOKUP(CC14,$A:$Z,2,FALSE)", "=VLOOKUP(CD14,$A:$Z,2,FALSE)"),
]

print("Testing formula adjustment function:\n")
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
