"""
Formula operations for Excel files.
Formulas are written DIRECTLY to Excel cells using openpyxl.

PERFORMANCE OPTIMIZED:
- Pre-compiled regex for cell reference adjustment
- Fast Python-native evaluation for common formulas (IF, OR, AND, LEFT, RIGHT, etc.)
- Falls back to 'formulas' library only for complex formulas (VLOOKUP, etc.)

NO win32com dependency!
"""

import openpyxl
import re
from pathlib import Path
from openpyxl.utils import column_index_from_string, get_column_letter
from core import wb_cache

# Pre-compile regex pattern for cell references (MUCH faster than compiling per-row)
CELL_REF_PATTERN = re.compile(r'(\$?)([A-Z]+)(\$?)(\d+)')

# Try to import formulas library for calculating formulas in Python
try:
    import formulas
    FORMULAS_AVAILABLE = True
except ImportError:
    FORMULAS_AVAILABLE = False


# ============================================================================
# FAST PYTHON-NATIVE FORMULA EVALUATION
# For common formulas like IF, OR, AND, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER
# AND VLOOKUP - using efficient dict-based lookup
# This is 100-1000x faster than the 'formulas' library for these cases
# ============================================================================

# Cache for VLOOKUP data to avoid re-reading the same range multiple times
_vlookup_cache = {}


def evaluate_formula_fast(formula_str, ws, row_num, wb=None):
    """
    Fast Python-native evaluation of common Excel formulas.
    
    Supports: IF, OR, AND, NOT, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER,
              CONCATENATE, VLOOKUP, IFERROR, &, =, <>, <, >, <=, >=, +, -, *, /
    
    Args:
        formula_str: The formula string (e.g., "=VLOOKUP(A2,B:C,2,FALSE)")
        ws: Current worksheet
        row_num: Current row number
        wb: Workbook (needed for cross-sheet VLOOKUP)
    
    Returns: (success: bool, value)
    """
    try:
        # Parse the formula and evaluate
        formula = formula_str.lstrip('=').strip()
        result = _eval_expression(formula, ws, row_num, wb)
        return True, result
    except Exception:
        return False, None


def _parse_range_reference(range_str, current_ws, wb=None):
    """
    Parse a range reference like 'holding!C:H' or 'A:B' or 'A1:D100'.
    
    Returns: (worksheet, start_col, end_col, start_row, end_row)
    Where start_row/end_row are None for full-column references like C:H
    """
    range_str = range_str.strip()
    
    # Check for sheet reference
    target_ws = current_ws
    if '!' in range_str:
        sheet_part, range_part = range_str.split('!', 1)
        sheet_name = sheet_part.strip().strip("'")
        if wb and sheet_name in wb.sheetnames:
            target_ws = wb[sheet_name]
        elif wb:
            # Try case-insensitive match
            for sn in wb.sheetnames:
                if sn.lower() == sheet_name.lower():
                    target_ws = wb[sn]
                    break
        range_str = range_part
    
    # Parse the range (C:H or A1:D100 or A:A)
    if ':' in range_str:
        start_ref, end_ref = range_str.split(':')
        
        # Check if it's a column-only reference (C:H)
        if start_ref.isalpha() and end_ref.isalpha():
            start_col = column_index_from_string(start_ref.upper())
            end_col = column_index_from_string(end_ref.upper())
            return target_ws, start_col, end_col, None, None
        
        # Parse cell references (A1:D100)
        start_match = CELL_REF_PATTERN.match(start_ref)
        end_match = CELL_REF_PATTERN.match(end_ref)
        
        if start_match and end_match:
            start_col = column_index_from_string(start_match.group(2))
            start_row = int(start_match.group(4))
            end_col = column_index_from_string(end_match.group(2))
            end_row = int(end_match.group(4))
            return target_ws, start_col, end_col, start_row, end_row
    
    return None, None, None, None, None


def _build_vlookup_index(target_ws, start_col, end_col, start_row, end_row):
    """
    Build a dict-based index for fast VLOOKUP.
    
    Returns: dict mapping lookup_value -> list of row data
    """
    cache_key = (id(target_ws), start_col, end_col, start_row, end_row)
    
    if cache_key in _vlookup_cache:
        return _vlookup_cache[cache_key]
    
    index = {}
    
    # Determine row range
    if start_row is None:
        start_row = 1
    if end_row is None:
        end_row = target_ws.max_row
    
    # Build index: key -> row data
    for row in range(start_row, end_row + 1):
        # Get the lookup key (first column in range)
        key_cell = target_ws.cell(row=row, column=start_col)
        key_value = key_cell.value
        
        # Skip empty keys
        if key_value is None or key_value == '':
            continue
        
        # Normalize key for comparison
        key_normalized = str(key_value).strip()
        
        # Store row data
        row_data = []
        for col in range(start_col, end_col + 1):
            cell_val = target_ws.cell(row=row, column=col).value
            row_data.append(cell_val)
        
        # Store first match (VLOOKUP returns first match)
        if key_normalized not in index:
            index[key_normalized] = row_data
    
    # Cache the index
    _vlookup_cache[cache_key] = index
    return index


def clear_vlookup_cache():
    """Clear the VLOOKUP cache (call between file operations)."""
    global _vlookup_cache
    _vlookup_cache = {}


def _get_cell_value(ws, col_letter, row_num):
    """Get cell value from worksheet."""
    try:
        cell = ws[f"{col_letter}{row_num}"]
        val = cell.value
        # If it's a formula, try to get cached value
        if isinstance(val, str) and val.startswith('='):
            return None  # Can't evaluate nested formulas in fast mode
        return val
    except Exception:
        return None


def _resolve_cell_ref(ref_str, ws, row_num):
    """Resolve a cell reference like A2, $A$2, etc. to its value."""
    ref_str = ref_str.strip()
    # IMPORTANT: require a FULL match; otherwise tokens like "$F2*100" would incorrectly match "$F2"
    match = CELL_REF_PATTERN.fullmatch(ref_str)
    if match:
        col_letter = match.group(2)
        ref_row = int(match.group(4))
        return _get_cell_value(ws, col_letter, ref_row)
    return None


def _parse_value(token, ws, row_num, wb=None):
    """Parse a token into a Python value."""
    token = token.strip()
    
    if not token:
        return None
    
    # String literal
    if (token.startswith('"') and token.endswith('"')) or (token.startswith("'") and token.endswith("'")):
        return token[1:-1]
    
    # Boolean
    if token.upper() == 'TRUE':
        return True
    if token.upper() == 'FALSE':
        return False
    
    # Number
    try:
        if '.' in token:
            return float(token)
        return int(token)
    except ValueError:
        pass
    
    # Cell reference (FULL match only)
    if CELL_REF_PATTERN.fullmatch(token):
        return _resolve_cell_ref(token, ws, row_num)
    
    # It might be an expression - try to evaluate it
    return _eval_expression(token, ws, row_num, wb)


def _split_args(args_str):
    """Split function arguments respecting nested parentheses and quotes."""
    args = []
    current = []
    depth = 0
    in_string = False
    quote_char = None
    
    for char in args_str:
        if char in '"\'':
            if not in_string:
                in_string = True
                quote_char = char
            elif char == quote_char:
                in_string = False
                quote_char = None
            current.append(char)
        elif char == '(' and not in_string:
            depth += 1
            current.append(char)
        elif char == ')' and not in_string:
            depth -= 1
            current.append(char)
        elif char == ',' and depth == 0 and not in_string:
            args.append(''.join(current).strip())
            current = []
        else:
            current.append(char)
    
    if current:
        args.append(''.join(current).strip())
    
    return args


def _eval_expression(expr, ws, row_num, wb=None):
    """Evaluate an expression."""
    expr = expr.strip()
    
    if not expr:
        return None
    
    # Check for function calls
    func_match = re.match(r'^([A-Z]+)\s*\((.*)\)$', expr, re.IGNORECASE | re.DOTALL)
    if func_match:
        func_name = func_match.group(1).upper()
        args_str = func_match.group(2)
        return _eval_function(func_name, args_str, ws, row_num, wb)
    
    # Check for comparison operators (handle these before arithmetic)
    # Order matters: check longer operators first
    for op in ['<>', '<=', '>=', '=', '<', '>']:
        # Find operator not inside quotes or parentheses
        pos = _find_operator(expr, op)
        if pos >= 0:
            left = expr[:pos].strip()
            right = expr[pos + len(op):].strip()
            left_val = _parse_value(left, ws, row_num, wb)
            right_val = _parse_value(right, ws, row_num, wb)
            
            # Handle string comparison
            if isinstance(left_val, str):
                left_val = left_val.strip()
            if isinstance(right_val, str):
                right_val = right_val.strip()
            
            if op == '=':
                return left_val == right_val
            elif op == '<>':
                return left_val != right_val
            elif op == '<':
                return left_val < right_val if left_val is not None and right_val is not None else False
            elif op == '>':
                return left_val > right_val if left_val is not None and right_val is not None else False
            elif op == '<=':
                return left_val <= right_val if left_val is not None and right_val is not None else False
            elif op == '>=':
                return left_val >= right_val if left_val is not None and right_val is not None else False
    
    # Check for concatenation with &
    pos = _find_operator(expr, '&')
    if pos >= 0:
        left = expr[:pos].strip()
        right = expr[pos + 1:].strip()
        left_val = _parse_value(left, ws, row_num, wb)
        right_val = _parse_value(right, ws, row_num, wb)
        return str(left_val or '') + str(right_val or '')

    # Arithmetic: handle * and / first, then + and -
    # We only split on operators at depth 0 (not inside parentheses/strings)
    for op in ['*', '/']:
        pos = _find_operator(expr, op)
        if pos >= 0:
            left = expr[:pos].strip()
            right = expr[pos + 1:].strip()
            a = _eval_expression(left, ws, row_num, wb)
            b = _eval_expression(right, ws, row_num, wb)

            # Excel treats blank as 0 in arithmetic
            if a is None or a == '':
                a = 0
            if b is None or b == '':
                b = 0

            try:
                a_num = float(a)
                b_num = float(b)
            except Exception:
                # If we can't coerce, return None to trigger fallback (safer)
                return None

            if op == '*':
                return a_num * b_num
            # division
            try:
                return a_num / b_num
            except ZeroDivisionError:
                return None

    # Addition/subtraction (lowest precedence of arithmetic)
    # NOTE: this is a simple implementation; unary minus is supported at the start (e.g., "-1")
    for op in ['+', '-']:
        pos = _find_operator(expr, op)
        if pos >= 0:
            # Unary minus: "-1" or "-A2"
            if op == '-' and pos == 0:
                v = _eval_expression(expr[1:].strip(), ws, row_num, wb)
                try:
                    return -float(v)
                except Exception:
                    return None

            left = expr[:pos].strip()
            right = expr[pos + 1:].strip()
            a = _eval_expression(left, ws, row_num, wb)
            b = _eval_expression(right, ws, row_num, wb)

            if a is None or a == '':
                a = 0
            if b is None or b == '':
                b = 0

            try:
                a_num = float(a)
                b_num = float(b)
            except Exception:
                return None

            if op == '+':
                return a_num + b_num
            return a_num - b_num
    
    # Simple value
    return _parse_value(expr, ws, row_num, wb)


def _find_operator(expr, op):
    """Find operator position not inside parentheses or quotes."""
    depth = 0
    in_string = False
    quote_char = None
    
    i = 0
    while i < len(expr):
        char = expr[i]
        
        if char in '"\'':
            if not in_string:
                in_string = True
                quote_char = char
            elif char == quote_char:
                in_string = False
                quote_char = None
        elif char == '(' and not in_string:
            depth += 1
        elif char == ')' and not in_string:
            depth -= 1
        elif depth == 0 and not in_string:
            if expr[i:i+len(op)] == op:
                return i
        
        i += 1
    
    return -1


def _eval_function(func_name, args_str, ws, row_num, wb=None):
    """Evaluate a function call."""
    args = _split_args(args_str)
    
    if func_name == 'IF':
        if len(args) >= 2:
            condition = _eval_expression(args[0], ws, row_num, wb)
            if condition:
                return _eval_expression(args[1], ws, row_num, wb)
            elif len(args) >= 3:
                return _eval_expression(args[2], ws, row_num, wb)
            else:
                return False
    
    elif func_name == 'OR':
        for arg in args:
            if _eval_expression(arg, ws, row_num, wb):
                return True
        return False
    
    elif func_name == 'AND':
        for arg in args:
            if not _eval_expression(arg, ws, row_num, wb):
                return False
        return True
    
    elif func_name == 'NOT':
        if args:
            return not _eval_expression(args[0], ws, row_num, wb)
        return True
    
    elif func_name == 'VLOOKUP':
        # VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
        # Fast Python-native VLOOKUP using dict-based index
        if len(args) >= 3:
            # Get lookup value
            lookup_value = _parse_value(args[0], ws, row_num, wb)
            if lookup_value is None:
                return None
            lookup_key = str(lookup_value).strip()
            
            # Parse table range
            table_range = args[1].strip()
            target_ws, start_col, end_col, start_row, end_row = _parse_range_reference(table_range, ws, wb)
            
            if target_ws is None:
                return None
            
            # Get column index (1-based)
            col_index = int(_parse_value(args[2], ws, row_num, wb))
            
            # Get range_lookup (optional, default FALSE for exact match)
            exact_match = True
            if len(args) >= 4:
                range_lookup = _parse_value(args[3], ws, row_num, wb)
                if range_lookup is True or (isinstance(range_lookup, str) and range_lookup.upper() == 'TRUE'):
                    exact_match = False
            
            # Build lookup index (cached for performance)
            lookup_index = _build_vlookup_index(target_ws, start_col, end_col, start_row, end_row)
            
            # Perform lookup
            if lookup_key in lookup_index:
                row_data = lookup_index[lookup_key]
                if col_index <= len(row_data):
                    return row_data[col_index - 1]
            
            # Not found - raise error (for IFERROR to catch)
            raise ValueError(f"VLOOKUP: {lookup_key} not found")
    
    elif func_name == 'LEFT':
        if len(args) >= 1:
            text = _parse_value(args[0], ws, row_num, wb)
            num_chars = int(_parse_value(args[1], ws, row_num, wb)) if len(args) >= 2 else 1
            if text is not None:
                return str(text)[:num_chars]
        return ''
    
    elif func_name == 'RIGHT':
        if len(args) >= 1:
            text = _parse_value(args[0], ws, row_num, wb)
            num_chars = int(_parse_value(args[1], ws, row_num, wb)) if len(args) >= 2 else 1
            if text is not None:
                return str(text)[-num_chars:] if num_chars > 0 else ''
        return ''
    
    elif func_name == 'MID':
        if len(args) >= 3:
            text = _parse_value(args[0], ws, row_num, wb)
            start = int(_parse_value(args[1], ws, row_num, wb))
            num_chars = int(_parse_value(args[2], ws, row_num, wb))
            if text is not None:
                return str(text)[start-1:start-1+num_chars]
        return ''
    
    elif func_name == 'LEN':
        if args:
            text = _parse_value(args[0], ws, row_num, wb)
            return len(str(text)) if text is not None else 0
        return 0

    elif func_name in ('FIND', 'SEARCH'):
        # FIND(find_text, within_text, [start_num])
        # FIND is case-sensitive; SEARCH is case-insensitive. Returns 1-based position.
        if len(args) >= 2:
            find_text  = str(_parse_value(args[0], ws, row_num, wb) or '')
            within     = str(_parse_value(args[1], ws, row_num, wb) or '')
            start_num  = int(_parse_value(args[2], ws, row_num, wb) or 1) if len(args) >= 3 else 1
            search_in  = within[start_num - 1:] if start_num > 1 else within
            if func_name == 'SEARCH':
                idx = search_in.lower().find(find_text.lower())
            else:
                idx = search_in.find(find_text)
            if idx == -1:
                return None   # #VALUE! in Excel — propagates as None → empty cell
            return idx + start_num   # convert back to 1-based absolute position
        return None

    elif func_name == 'SUBSTITUTE':
        # SUBSTITUTE(text, old_text, new_text, [instance_num])
        if len(args) >= 3:
            text      = str(_parse_value(args[0], ws, row_num, wb) or '')
            old_text  = str(_parse_value(args[1], ws, row_num, wb) or '')
            new_text  = str(_parse_value(args[2], ws, row_num, wb) or '')
            if len(args) >= 4:
                instance = int(_parse_value(args[3], ws, row_num, wb) or 1)
                count = 0
                result = ''
                i = 0
                while i < len(text):
                    pos = text.find(old_text, i)
                    if pos == -1:
                        result += text[i:]
                        break
                    count += 1
                    if count == instance:
                        result += text[i:pos] + new_text
                        i = pos + len(old_text)
                        result += text[i:]
                        break
                    result += text[i:pos + len(old_text)]
                    i = pos + len(old_text)
                return result
            return text.replace(old_text, new_text) if old_text else text
        return ''

    elif func_name == 'REPLACE':
        # REPLACE(old_text, start_num, num_chars, new_text)
        if len(args) >= 4:
            text      = str(_parse_value(args[0], ws, row_num, wb) or '')
            start     = int(_parse_value(args[1], ws, row_num, wb) or 1)
            num_chars = int(_parse_value(args[2], ws, row_num, wb) or 0)
            new_text  = str(_parse_value(args[3], ws, row_num, wb) or '')
            return text[:start - 1] + new_text + text[start - 1 + num_chars:]
        return ''

    elif func_name == 'TRIM':
        if args:
            text = _parse_value(args[0], ws, row_num, wb)
            return str(text).strip() if text is not None else ''
        return ''
    
    elif func_name == 'UPPER':
        if args:
            text = _parse_value(args[0], ws, row_num, wb)
            return str(text).upper() if text is not None else ''
        return ''
    
    elif func_name == 'LOWER':
        if args:
            text = _parse_value(args[0], ws, row_num, wb)
            return str(text).lower() if text is not None else ''
        return ''
    
    elif func_name == 'CONCATENATE':
        result = ''
        for arg in args:
            val = _parse_value(arg, ws, row_num, wb)
            result += str(val) if val is not None else ''
        return result
    
    elif func_name == 'IFERROR':
        if len(args) >= 2:
            try:
                result = _eval_expression(args[0], ws, row_num, wb)
                if result is None:
                    return _eval_expression(args[1], ws, row_num, wb)
                return result
            except Exception:
                return _eval_expression(args[1], ws, row_num, wb)
        return None
    
    elif func_name == 'ISBLANK':
        if args:
            val = _parse_value(args[0], ws, row_num, wb)
            return val is None or val == ''
        return True
    
    elif func_name == 'TEXT':
        if len(args) >= 1:
            val = _parse_value(args[0], ws, row_num, wb)
            return str(val) if val is not None else ''
        return ''

    elif func_name == 'TEXTBEFORE':
        # TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
        # Returns text before the delimiter. instance_num defaults to 1 (first occurrence).
        if len(args) >= 2:
            text = str(_parse_value(args[0], ws, row_num, wb) or '')
            delim = str(_parse_value(args[1], ws, row_num, wb) or '')
            instance = int(_parse_value(args[2], ws, row_num, wb) or 1) if len(args) >= 3 else 1
            if not delim:
                return text
            parts = text.split(delim)
            idx = instance - 1 if instance > 0 else len(parts) + instance - 1
            if 0 <= idx < len(parts) - 1:
                return delim.join(parts[:idx + 1]) if idx > 0 else parts[0]
            # delimiter not found — return if_not_found or raise
            if len(args) >= 6:
                return str(_parse_value(args[5], ws, row_num, wb) or '')
            return ''
        return ''

    elif func_name == 'TEXTAFTER':
        # TEXTAFTER(text, delimiter, [instance_num], ...)
        # Returns text after the delimiter.
        if len(args) >= 2:
            text = str(_parse_value(args[0], ws, row_num, wb) or '')
            delim = str(_parse_value(args[1], ws, row_num, wb) or '')
            instance = int(_parse_value(args[2], ws, row_num, wb) or 1) if len(args) >= 3 else 1
            if not delim:
                return text
            parts = text.split(delim)
            idx = instance - 1 if instance > 0 else len(parts) + instance - 1
            if 0 <= idx < len(parts) - 1:
                return delim.join(parts[idx + 1:])
            if len(args) >= 6:
                return str(_parse_value(args[5], ws, row_num, wb) or '')
            return ''
        return ''

    elif func_name == 'TEXTJOIN':
        # TEXTJOIN(delimiter, ignore_empty, text1, ...)
        if len(args) >= 3:
            delim = str(_parse_value(args[0], ws, row_num, wb) or '')
            ignore_empty = _parse_value(args[1], ws, row_num, wb)
            parts = []
            for a in args[2:]:
                v = _parse_value(a, ws, row_num, wb)
                s = str(v) if v is not None else ''
                if ignore_empty and not s:
                    continue
                parts.append(s)
            return delim.join(parts)
        return ''

    # Unsupported function - return None to trigger fallback
    return None


def can_evaluate_fast(formula_str):
    """
    Check if a formula can be evaluated with the fast Python-native engine.
    Now supports VLOOKUP with fast dict-based index!
    Returns True for: IF, OR, AND, LEFT, RIGHT, VLOOKUP, IFERROR, etc.
    Returns False for: HLOOKUP, INDEX/MATCH, SUMIF, INDIRECT, etc.
    """
    formula_upper = formula_str.upper()
    
    # Functions that require the slow library (VLOOKUP is now fast!)
    # Use regex to match function names followed by ( to avoid substring matches
    # e.g., "LOOKUP" shouldn't match inside "VLOOKUP"
    complex_functions = [
        'HLOOKUP', 'INDEX', 'MATCH', 'INDIRECT',
        'SUMIF', 'SUMIFS', 'COUNTIF', 'COUNTIFS', 'AVERAGEIF',
        'OFFSET', 'CHOOSE', 'ROW', 'COLUMN', 'ADDRESS'
    ]
    # TEXTBEFORE, TEXTAFTER, TEXTJOIN are natively supported by the fast engine
    # (no exclusion needed — they fall through to the evaluator above)
    
    # Special check for LOOKUP that isn't part of VLOOKUP/HLOOKUP
    # Match LOOKUP( but not VLOOKUP( or HLOOKUP(
    if re.search(r'(?<!V)(?<!H)LOOKUP\s*\(', formula_upper):
        return False
    
    for func in complex_functions:
        # Check for function name followed by (
        pattern = func + r'\s*\('
        if re.search(pattern, formula_upper):
            return False
    
    return True


def calculate_workbook_formulas(file_path):
    """
    Calculate all formulas in an Excel workbook using the 'formulas' library.
    Returns a dict mapping cell references to calculated values.
    """
    if not FORMULAS_AVAILABLE:
        return None
    
    try:
        # Load the workbook with formulas library
        xl_model = formulas.ExcelModel().loads(str(file_path)).finish()
        # Calculate all formulas
        solution = xl_model.calculate()
        return solution
    except Exception:
        return None


def adjust_formula_for_row_fast(formula_str, row_diff):
    """
    Fast formula row adjustment using pre-compiled regex.
    
    Args:
        formula_str: Excel formula string
        row_diff: Difference between target and original row
    
    Returns:
        Adjusted formula string
    """
    if row_diff == 0:
        return formula_str
    
    def replacer(match):
        col_abs = match.group(1)  # $ before column
        col_letter = match.group(2)  # Column letter
        row_abs = match.group(3)  # $ before row
        row_num = int(match.group(4))  # Row number
        
        # Only adjust row if it's not absolute ($)
        if row_abs == '$':
            return match.group(0)  # Return unchanged
        else:
            new_row = row_num + row_diff
            return f"{col_abs}{col_letter}{row_abs}{new_row}"
    
    return CELL_REF_PATTERN.sub(replacer, formula_str)


def run_formula_step(file_path, sheet_name, formula, column_letter, start_row, end_row, convert_to_values=False):
    """
    Apply a formula to a column range in an Excel file.
    Formula is written DIRECTLY to Excel cells - Excel evaluates it when the file is opened.
    
    PERFORMANCE OPTIMIZED:
    - Pre-compiled regex for cell references
    - Single workbook load/save
    - Batch cell writes
    
    Args:
        file_path: Path to Excel file (.xlsx)
        sheet_name: Sheet name
        formula: Excel formula (must start with =)
        column_letter: Column letter (e.g., 'A', 'AB')
        start_row: Starting row number
        end_row: Ending row number or 'last'
        convert_to_values: If True, calculate and replace formulas with values
    
    Returns:
        Tuple of (success: bool, message: str)
    """
    if not formula or not formula.lstrip().startswith('=') or not column_letter:
        return False, "Missing or invalid fields (Formula must start with =, Column required)"
    
    formula_clean = formula.lstrip()
    column_letter = column_letter.upper()
    
    try:
        # Load workbook once
        wb = wb_cache.load(file_path)
        
        # Get worksheet
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            actual_sheet = sheet_name
        else:
            ws = wb.active
            actual_sheet = ws.title
        
        # Determine end row
        if end_row == "last" or end_row == "":
            end_row_num = ws.max_row
        else:
            end_row_num = int(end_row)
        
        start_row_num = int(start_row)
        
        # Get column index for direct cell access (faster than string concat)
        col_idx = column_index_from_string(column_letter)
        
        # Pre-calculate the base row from the formula (extract first row number in formula)
        base_row_match = CELL_REF_PATTERN.search(formula_clean)
        if base_row_match:
            formula_base_row = int(base_row_match.group(4))
        else:
            formula_base_row = start_row_num
        
        # Write formulas to cells in batch
        for row in range(start_row_num, end_row_num + 1):
            # Calculate row difference from formula's base row
            row_diff = row - formula_base_row
            
            # Adjust formula using fast pre-compiled regex
            adjusted_formula = adjust_formula_for_row_fast(formula_clean, row_diff)
            
            # Direct cell access by row/col index (faster)
            ws.cell(row=row, column=col_idx).value = adjusted_formula
        
        # Save workbook with formulas
        wb_cache.save(wb, file_path)
        wb.close()
        
        total_rows = end_row_num - start_row_num + 1
        
        # If convert_to_values, calculate formulas and replace with values
        if convert_to_values:
            file_name = Path(file_path).name
            
            # FAST PATH: Try Python-native evaluation for simple formulas
            # This is 100-1000x faster than the 'formulas' library
            if can_evaluate_fast(formula_clean):
                try:
                    wb = wb_cache.load(file_path)
                    ws = wb[actual_sheet]

                    # Clear VLOOKUP cache for fresh data
                    clear_vlookup_cache()
                    
                    converted_count = 0
                    for row in range(start_row_num, end_row_num + 1):
                        # Get the formula for this row
                        row_diff = row - formula_base_row
                        adjusted_formula = adjust_formula_for_row_fast(formula_clean, row_diff)
                        
                        # Evaluate using fast Python-native engine (pass wb for cross-sheet VLOOKUP)
                        success, value = evaluate_formula_fast(adjusted_formula, ws, row, wb)
                        
                        if success and value is not None:
                            ws.cell(row=row, column=col_idx).value = value
                            converted_count += 1
                    
                    wb_cache.save(wb, file_path)
                    wb.close()

                    if converted_count > 0:
                        return True, f"Formula applied to {column_letter}{start_row_num}:{column_letter}{end_row_num} ({total_rows} rows, fast-converted: {converted_count} cells)"
                    
                except Exception:
                    pass  # Fall through to slow method
            
            # SLOW PATH: Use 'formulas' library for complex formulas (VLOOKUP, etc.)
            if FORMULAS_AVAILABLE:
                try:
                    # Calculate all formulas in the workbook
                    xl_model = formulas.ExcelModel().loads(str(file_path)).finish()
                    solution = xl_model.calculate()
                    
                    # Re-open workbook to replace formulas with calculated values
                    wb = wb_cache.load(file_path)
                    ws = wb[actual_sheet]

                    converted_count = 0
                    for row in range(start_row_num, end_row_num + 1):
                        # Try multiple reference formats that formulas library might use
                        cell_refs = [
                            f"'[{file_name}]{actual_sheet}'!{column_letter}{row}",
                            f"'[{file_name}]{actual_sheet.upper()}'!{column_letter}{row}",
                            f"[{file_name}]{actual_sheet}!{column_letter}{row}",
                            f"[{file_name}]{actual_sheet.upper()}!{column_letter}{row}",
                            f"'{actual_sheet}'!{column_letter}{row}",
                            f"'{actual_sheet.upper()}'!{column_letter}{row}",
                            f"{actual_sheet}!{column_letter}{row}",
                            f"{actual_sheet.upper()}!{column_letter}{row}",
                        ]
                        
                        value = None
                        for ref in cell_refs:
                            if ref in solution:
                                raw_value = solution[ref]
                                if hasattr(raw_value, 'value'):
                                    value = raw_value.value
                                elif hasattr(raw_value, 'tolist'):
                                    value = raw_value.tolist()
                                else:
                                    value = raw_value
                                
                                if hasattr(value, 'tolist'):
                                    value = value.tolist()
                                
                                while isinstance(value, (list, tuple)) and len(value) > 0:
                                    value = value[0]
                                
                                if hasattr(value, 'item'):
                                    value = value.item()
                                break
                        
                        if value is not None:
                            ws.cell(row=row, column=col_idx).value = value
                            converted_count += 1
                    
                    wb_cache.save(wb, file_path)
                    wb.close()

                    if converted_count > 0:
                        return True, f"Formula applied to {column_letter}{start_row_num}:{column_letter}{end_row_num} ({total_rows} rows, calculated & converted: {converted_count} cells)"
                    
                except Exception:
                    pass
            
            # Fallback: Try cached values from Excel
            wb_data = wb_cache.load(file_path, data_only=True)
            ws_data = wb_data[actual_sheet] if actual_sheet in wb_data.sheetnames else wb_data.active

            wb = wb_cache.load(file_path)
            ws = wb[actual_sheet] if actual_sheet in wb.sheetnames else wb.active
            
            converted_count = 0
            none_count = 0
            for row in range(start_row_num, end_row_num + 1):
                cell_data = ws_data.cell(row=row, column=col_idx)
                
                if cell_data.value is not None:
                    ws.cell(row=row, column=col_idx).value = cell_data.value
                    converted_count += 1
                else:
                    none_count += 1
            
            wb_data.close()
            wb_cache.save(wb, file_path)
            wb.close()

            if none_count > 0 and converted_count == 0:
                if not FORMULAS_AVAILABLE:
                    return True, f"Formula applied to {column_letter}{start_row_num}:{column_letter}{end_row_num} ({total_rows} rows). ⚠️ Install 'formulas' library for auto-calculation, or open in Excel first."
                return True, f"Formula applied to {column_letter}{start_row_num}:{column_letter}{end_row_num} ({total_rows} rows). ⚠️ Could not calculate. Open in Excel first."
            elif none_count > 0:
                return True, f"Formula applied ({total_rows} rows). Converted {converted_count} cells ({none_count} had no value)"
            else:
                return True, f"Formula applied to {column_letter}{start_row_num}:{column_letter}{end_row_num} ({total_rows} rows, converted to values)"
        
        msg = f"Formula applied to {column_letter}{start_row_num}:{column_letter}{end_row_num} ({total_rows} rows)"
        return True, msg
        
    except Exception as e:
        return False, f"Error applying formula: {str(e)}"


def adjust_formula_for_row(formula_str, original_row, target_row):
    """
    Adjust cell references in a formula for a different row.
    (Legacy function - kept for compatibility, uses fast version internally)
    """
    return adjust_formula_for_row_fast(formula_str, target_row - original_row)


def apply_multiple_formulas_to_xls(file_path, sheet_name, formulas_batch):
    """
    Apply multiple formulas to an Excel file in one operation.
    PERFORMANCE OPTIMIZED: Pre-compiled regex, column index caching
    
    Args:
        file_path: Path to Excel file
        sheet_name: Sheet name
        formulas_batch: List of dicts with keys:
            - formula: Formula string
            - column_letter: Column letter
            - start_row: Start row
            - end_row: End row
            - convert_to_values: Boolean
    
    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        wb = wb_cache.load(file_path)

        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active

        total_cells = 0
        
        # Apply all formulas
        for formula_spec in formulas_batch:
            formula = formula_spec.get('formula', '').lstrip()
            column_letter = formula_spec.get('column_letter', '').upper()
            start_row = int(formula_spec.get('start_row', 2))
            end_row_str = formula_spec.get('end_row', 'last')
            
            if not formula or not formula.startswith('=') or not column_letter:
                continue
            
            # Get column index (faster than string concat)
            col_idx = column_index_from_string(column_letter)
            
            # Determine end row
            if end_row_str == "last":
                end_row = ws.max_row
            else:
                end_row = int(end_row_str)
            
            # Get base row from formula
            base_row_match = CELL_REF_PATTERN.search(formula)
            if base_row_match:
                formula_base_row = int(base_row_match.group(4))
            else:
                formula_base_row = start_row
            
            # Write formulas using fast method
            for row in range(start_row, end_row + 1):
                row_diff = row - formula_base_row
                adjusted_formula = adjust_formula_for_row_fast(formula, row_diff)
                ws.cell(row=row, column=col_idx).value = adjusted_formula
                total_cells += 1
        
        # Save
        wb_cache.save(wb, file_path)
        wb.close()

        return True, f"Applied {len(formulas_batch)} formulas ({total_cells} cells) successfully"
        
    except Exception as e:
        return False, f"Error applying multiple formulas: {str(e)}"


def run_convert_to_values_step(file_path, sheet_name, col_letter, start_row, end_row):
    """
    Convert formulas in a column range to their calculated values.
    
    PERFORMANCE OPTIMIZED:
    - Uses fast Python-native evaluation for simple formulas (IF, OR, LEFT, etc.)
    - Falls back to 'formulas' library for complex formulas (VLOOKUP, etc.)
    - Falls back to cached values if neither works
    
    Args:
        file_path: Path to Excel file
        sheet_name: Sheet name
        col_letter: Column letter
        start_row: Start row (can be string or int)
        end_row: End row or 'last'
    
    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        col_letter = col_letter.upper()
        col_idx = column_index_from_string(col_letter)
        
        # Determine row range and actual sheet name first
        wb = wb_cache.load(file_path)
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            actual_sheet = sheet_name
        else:
            ws = wb.active
            actual_sheet = ws.title
        
        start_row_num = int(start_row) if str(start_row).isdigit() else 2
        if end_row == "last" or end_row == "":
            end_row_num = ws.max_row
        else:
            end_row_num = int(end_row)
        
        total_rows = end_row_num - start_row_num + 1
        file_name = Path(file_path).name
        
        # FAST PATH: Try Python-native evaluation for simple formulas
        # Check first cell to see if we can use fast evaluation
        first_cell = ws.cell(row=start_row_num, column=col_idx)
        first_formula = first_cell.value if isinstance(first_cell.value, str) and first_cell.value.startswith('=') else None
        
        if first_formula and can_evaluate_fast(first_formula):
            # Clear VLOOKUP cache for fresh data
            clear_vlookup_cache()
            
            converted_count = 0
            for row in range(start_row_num, end_row_num + 1):
                cell = ws.cell(row=row, column=col_idx)
                cell_value = cell.value
                
                if isinstance(cell_value, str) and cell_value.startswith('='):
                    # Evaluate using fast Python-native engine (pass wb for cross-sheet VLOOKUP)
                    success, value = evaluate_formula_fast(cell_value, ws, row, wb)
                    
                    if success and value is not None:
                        cell.value = value
                        converted_count += 1
            
            if converted_count > 0:
                wb_cache.save(wb, file_path)
                wb.close()
                return True, f"Fast-converted {converted_count}/{total_rows} cells to values in {col_letter}{start_row_num}:{col_letter}{end_row_num}"
        
        wb.close()
        
        # SLOW PATH: Use 'formulas' library for complex formulas (VLOOKUP, etc.)
        if FORMULAS_AVAILABLE:
            try:
                xl_model = formulas.ExcelModel().loads(str(file_path)).finish()
                solution = xl_model.calculate()
                
                wb = wb_cache.load(file_path)
                ws = wb[actual_sheet]

                converted_count = 0
                for row in range(start_row_num, end_row_num + 1):
                    # Try multiple reference formats that formulas library might use
                    cell_refs = [
                        f"'[{file_name}]{actual_sheet}'!{col_letter}{row}",
                        f"'[{file_name}]{actual_sheet.upper()}'!{col_letter}{row}",
                        f"[{file_name}]{actual_sheet}!{col_letter}{row}",
                        f"[{file_name}]{actual_sheet.upper()}!{col_letter}{row}",
                        f"'{actual_sheet}'!{col_letter}{row}",
                        f"'{actual_sheet.upper()}'!{col_letter}{row}",
                        f"{actual_sheet}!{col_letter}{row}",
                        f"{actual_sheet.upper()}!{col_letter}{row}",
                    ]
                    
                    value = None
                    for ref in cell_refs:
                        if ref in solution:
                            raw_value = solution[ref]
                            if hasattr(raw_value, 'value'):
                                value = raw_value.value
                            elif hasattr(raw_value, 'tolist'):
                                value = raw_value.tolist()
                            else:
                                value = raw_value
                            
                            if hasattr(value, 'tolist'):
                                value = value.tolist()
                            
                            while isinstance(value, (list, tuple)) and len(value) > 0:
                                value = value[0]
                            
                            if hasattr(value, 'item'):
                                value = value.item()
                            break
                    
                    if value is not None:
                        ws.cell(row=row, column=col_idx).value = value
                        converted_count += 1
                
                wb_cache.save(wb, file_path)
                wb.close()

                if converted_count > 0:
                    return True, f"Calculated & converted {converted_count}/{total_rows} cells to values in {col_letter}{start_row_num}:{col_letter}{end_row_num}"
                
            except Exception:
                pass  # Fall through to cached values method
        
        # Fallback: Use cached values from Excel
        wb = wb_cache.load(file_path)
        ws = wb[actual_sheet]

        wb_data = wb_cache.load(file_path, data_only=True)
        ws_data = wb_data[actual_sheet] if actual_sheet in wb_data.sheetnames else wb_data.active
        
        converted_count = 0
        none_count = 0
        for row in range(start_row_num, end_row_num + 1):
            cell_data = ws_data.cell(row=row, column=col_idx)
            
            if cell_data.value is not None:
                ws.cell(row=row, column=col_idx).value = cell_data.value
                converted_count += 1
            else:
                none_count += 1
        
        wb_data.close()
        wb_cache.save(wb, file_path)
        wb.close()
        
        if none_count > 0 and converted_count == 0:
            if not FORMULAS_AVAILABLE:
                return True, f"⚠️ Install 'formulas' library for auto-calculation: pip install formulas"
            return True, f"⚠️ Could not calculate formulas in {col_letter}{start_row_num}:{col_letter}{end_row_num}."
        elif none_count > 0:
            return True, f"Converted {converted_count}/{total_rows} cells ({none_count} could not be calculated)"
        else:
            return True, f"Converted formulas to values in {col_letter}{start_row_num}:{col_letter}{end_row_num} ({total_rows} rows)"
        
    except Exception as e:
        return False, f"Error converting to values: {str(e)}"


