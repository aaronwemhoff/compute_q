# utils/parser.py
# ---------------------------------
# Utilities to read the Excel file and extract:
#  - The platform names (rows)
#  - The metric names (columns)
#  - The numeric values per (platform, metric)
#  - The comparison operator per (platform, metric) derived from cell color
#
# Beginner notes:
# - We use openpyxl to access cell colors and values directly.
# - Pandas alone doesn't keep cell fill colors when reading Excel.
# ---------------------------------

from typing import Dict, Any, List
import pandas as pd
from openpyxl import load_workbook

# Heuristics to detect red/green shades.
# We look for common hex substrings in the cell.fill.fgColor/start_color fields.
GREEN_HINTS = {"00FF00", "00CC00", "33CC33", "66FF66", "009900"}
RED_HINTS   = {"FF0000", "CC0000", "CC3333", "990000", "FF6666"}

def _read_rgb(cell) -> str:
    """
    Try to get an uppercase 6- or 8-hex color code string from an Excel cell fill.
    Returns '' if none found.
    """
    try:
        # openpyxl may store color in different places depending on theme/format
        color = None
        if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
            color = cell.fill.fgColor.rgb  # often ARGB like 'FF00FF00'
        elif cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
            color = cell.fill.start_color.rgb
        if not color:
            return ""
        color = color.upper()
        # Normalize to 6-hex if possible (strip alpha FF prefix like 'FFRRGGBB')
        if len(color) == 8:
            color = color[2:]  # drop the alpha
        # Ensure we only return hex digits
        return "".join([c for c in color if c in "0123456789ABCDEF"])
    except Exception:
        return ""

def _operator_from_color(rgb: str) -> str:
    """
    Map a cell fill color to a comparison operator key.
    - 'le' => user_value <= platform_value (green cells)
    - 'gt' => user_value >  platform_value (red cells)
    - default 'ge' => user_value >= platform_value (uncolored/unknown)
    """
    rgb = (rgb or "").upper()
    # If we have an 8-hex, drop leading alpha
    if len(rgb) == 8:
        rgb = rgb[2:]

    # Check hints
    for g in GREEN_HINTS:
        if g in rgb:
            return "le"
    for r in RED_HINTS:
        if r in rgb:
            return "gt"
    # Fallback
    return "ge"

def detect_operator_label(op_key: str) -> str:
    """
    Convert internal operator key to a friendly label for help text.
    """
    mapping = {"le": "≤ (user ≤ platform value)",
               "gt": "> (user > platform value)",
               "ge": "≥ (user ≥ platform value)"}
    return mapping.get(op_key, "≥ (user ≥ platform value)")

def compare_value(user_val: float, platform_val: float, op_key: str) -> bool:
    """
    Compare two numbers using the operator indicated by op_key.
    """
    try:
        u = float(user_val)
        p = float(platform_val)
    except Exception:
        return False

    if op_key == "le":
        return u <= p
    elif op_key == "gt":
        return u > p
    else:  # 'ge' default
        return u >= p

def load_questionnaire(xlsx_path, header_row: int = 4, start_row: int = 5, end_row: int = 25) -> Dict[str, Any]:
    """
    Read the Excel workbook and extract platform/metric/value/operator information.

    Args:
        xlsx_path: Path or file-like object for the .xlsx
        header_row: 1-indexed row number containing column headers (metric names). First header is platform label.
        start_row: 1-indexed first data row (platform rows start here)
        end_row:   1-indexed last data row (inclusive)

    Returns:
        {
          'platforms_df': pd.DataFrame (index=platform names, columns=metric names, values=numbers),
          'metrics': List[str],
          'ops_grid': {platform: {metric: 'le'|'gt'|'ge'}},
          'values_grid': {platform: {metric: float}},
        }
    """
    wb = load_workbook(filename=xlsx_path, data_only=True)
    ws = wb.active  # use the first sheet by default

    # Read headers from the specified row
    headers = []
    max_col = ws.max_column
    for c in range(1, max_col + 1):
        headers.append(ws.cell(row=header_row, column=c).value)

    # The first header is the platform name label; remaining are metric names
    if not headers or len(headers) < 2:
        raise ValueError("Expected at least 2 columns (Platform + at least one metric).")

    platform_header = headers[0] or "Platform"
    metric_headers = [h if h is not None else f"Metric {i}" for i, h in enumerate(headers[1:], start=1)]

    # Loop over rows to collect data
    platforms = []
    rows_data = []
    ops_grid = {}      # {platform: {metric: op}}
    values_grid = {}   # {platform: {metric: value}}

    for r in range(start_row, end_row + 1):
        name_cell = ws.cell(row=r, column=1)
        platform_name = name_cell.value

        # Skip empty platform rows
        if platform_name is None or str(platform_name).strip() == "":
            continue

        row_vals = []
        row_ops = {}
        row_vals_map = {}

        for c in range(2, max_col + 1):
            metric = metric_headers[c - 2]  # since metrics start at col 2
            cell = ws.cell(row=r, column=c)
            raw_val = cell.value

            # Convert obvious non-numeric to NaN (we expect numbers for scoring)
            try:
                val = float(raw_val)
            except Exception:
                val = float('nan')

            rgb = _read_rgb(cell)
            op = _operator_from_color(rgb)

            row_vals.append(val)
            row_ops[metric] = op
            row_vals_map[metric] = val

        platforms.append(platform_name)
        rows_data.append(row_vals)
        ops_grid[platform_name] = row_ops
        values_grid[platform_name] = row_vals_map

    # Create DataFrame with platforms as rows and metrics as columns
    df = pd.DataFrame(rows_data, index=platforms, columns=metric_headers)

    return {
        'platforms_df': df,
        'metrics': metric_headers,
        'ops_grid': ops_grid,
        'values_grid': values_grid
    }
