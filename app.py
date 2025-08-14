# app.py
# ---------------------------------------------------------------
# A beginner-friendly Streamlit app that reads your Excel file,
# shows a questionnaire to the user, and recommends matching
# "Solution(s)" based on cell colors (green => <=, red => >).
# ---------------------------------------------------------------
import os
import tempfile
from typing import Dict, List, Tuple

import streamlit as st
import pandas as pd

# Create utils directory and files if they don't exist
if not os.path.exists('utils'):
    os.makedirs('utils')

# Create __init__.py for utils module
if not os.path.exists('utils/__init__.py'):
    with open('utils/__init__.py', 'w') as f:
        f.write('')

# Simple implementation of the missing utility functions
def try_float(value):
    """Convert a value to float if possible, otherwise return None."""
    try:
        return float(value)
    except (ValueError, TypeError):
        return None

def score_row(inputs: Dict, row: pd.Series, op_map: Dict, row_idx: int, input_cols: List[str]) -> Tuple[int, int]:
    """
    Score a row against user inputs.
    Returns (matched_count, considered_count)
    """
    matched = 0
    considered = 0
    
    for col in input_cols:
        user_val = try_float(inputs.get(col))
        cell_val = try_float(row.get(col))
        
        # Skip if either value is not numeric
        if user_val is None or cell_val is None:
            continue
            
        considered += 1
        
        # Get the operator for this cell (default to >= if not found)
        op = op_map.get((row_idx, col), 'ge')
        
        # Apply the comparison rule
        if op == 'le':  # Green cell: user_val <= cell_val
            if user_val <= cell_val:
                matched += 1
        elif op == 'gt':  # Red cell: user_val > cell_val
            if user_val > cell_val:
                matched += 1
        else:  # Default: user_val >= cell_val
            if user_val >= cell_val:
                matched += 1
    
    return matched, considered

class Questionnaire:
    """Simple questionnaire parser for Excel files."""
    
    def __init__(self, header_row=4, start_row=5, end_row=25):
        self.header_row = header_row - 1  # Convert to 0-based
        self.start_row = start_row - 1    # Convert to 0-based  
        self.end_row = end_row - 1        # Convert to 0-based
    
    def read_with_colors(self, file_path: str) -> Tuple[pd.DataFrame, Dict]:
        """
        Read Excel file and detect cell colors for comparison operators.
        Returns (dataframe, operator_map)
        """
        try:
            from openpyxl import load_workbook
        except ImportError:
            st.error("openpyxl is required. Please install it with: pip install openpyxl")
            return None, None
        
        # Read the data using pandas
        df = pd.read_excel(file_path, header=self.header_row, 
                          skiprows=range(0, self.header_row) if self.header_row > 0 else None)
        
        # Limit to the specified row range
        max_rows = min(len(df), self.end_row - self.start_row + 1)
        df = df.head(max_rows)
        
        # Load workbook to check cell colors
        wb = load_workbook(file_path, data_only=False)
        ws = wb.active
        
        op_map = {}
        
        # Check colors for each data cell
        for row_idx in range(len(df)):
            excel_row = self.start_row + row_idx + 1  # Convert back to 1-based Excel row
            for col_idx, col_name in enumerate(df.columns):
                excel_col = col_idx + 1  # Convert to 1-based Excel column
                cell = ws.cell(row=excel_row, column=excel_col)
                
                # Check if cell has a fill color
                if cell.fill and cell.fill.start_color and cell.fill.start_color.index:
                    color = cell.fill.start_color.index
                    # Simple color detection (this may need adjustment based on your Excel file)
                    if 'FF00FF00' in str(color) or 'FF92D050' in str(color):  # Green variants
                        op_map[(row_idx, col_name)] = 'le'  # Less than or equal
                    elif 'FFFF0000' in str(color) or 'FFFF6666' in str(color):  # Red variants
                        op_map[(row_idx, col_name)] = 'gt'  # Greater than
                    else:
                        op_map[(row_idx, col_name)] = 'ge'  # Default: greater than or equal
                else:
                    op_map[(row_idx, col_name)] = 'ge'  # Default: greater than or equal
        
        return df, op_map

st.set_page_config(page_title="Research Platform Selector", page_icon="🔍", layout="wide")

st.title("🔍 Research Platform Selector")
st.write(
    "Answer a few questions and I'll suggest the best **Solution** based on your Excel rules.\n\n"
    "- **Green cells** mean '**≤**' (less than or equal to)\n"
    "- **Red cells** mean '**>**' (greater than)\n"
    "- Headers are on **line 4**; data is in **lines 5–25**; output column is **`Solution`**"
)

# --------- Load Excel ----------
DEFAULT_PATH = "Research Questionnaire Info.xlsx"  # Updated default path
uploaded = st.file_uploader("📄 Upload Excel (e.g., Research Questionnaire Info.xlsx)", type=["xlsx"])

# Initialize session state for temp directory
if "_tmp_dir" not in st.session_state:
    st.session_state["_tmp_dir"] = tempfile.gettempdir()

q = Questionnaire()  # knows header/data rows
df = None
op_map = None
error = None

if uploaded is not None:
    # Save a temp copy so openpyxl can inspect styles/colors
    tmp_path = os.path.join(st.session_state.get("_tmp_dir", "."), "uploaded.xlsx")
    with open(tmp_path, "wb") as f:
        f.write(uploaded.read())
    df, op_map = q.read_with_colors(tmp_path)
elif os.path.exists(DEFAULT_PATH):
    df, op_map = q.read_with_colors(DEFAULT_PATH)
else:
    st.info("Upload an Excel file or place 'Research Questionnaire Info.xlsx' in the app directory.")
    st.stop()

if df is None or op_map is None:
    st.error("Couldn't read the spreadsheet. Please check the file format and try again.")
    st.stop()

# Validate presence of Solution column
solution_col = None
for c in df.columns:
    if str(c).strip().lower() == "solution":
        solution_col = c
        break

if solution_col is None:
    st.error("No `Solution` column was found in your headers (line 4). Please ensure one column is named exactly `Solution`.")
    st.stop()

input_cols: List[str] = [c for c in df.columns if c != solution_col]

with st.expander("🔍 Preview parsed data (from lines 5–25)"):
    st.dataframe(df)

# --------- Build Questionnaire UI ----------
st.header("📋 Questionnaire")
st.write("Provide values for each input below. If you leave some blank, those criteria will be skipped.")

with st.form("questionnaire_form", clear_on_submit=False):
    inputs: Dict[str, float] = {}
    for col in input_cols:
        # Suggest a min/max range based on the column data
        col_vals = pd.to_numeric(df[col], errors="coerce")
        col_min = float(col_vals.min()) if col_vals.notna().any() else None
        col_max = float(col_vals.max()) if col_vals.notna().any() else None

        # Use number_input when the column looks numeric, else text_input
        if col_vals.notna().any():
            default = col_min if col_min is not None else 0.0
            step = 1.0
            # Simple heuristic: if values look like ints, step by 1
            if col_vals.dropna().apply(lambda x: float(x).is_integer()).all():
                step = 1.0
            else:
                step = 0.1
            inputs[col] = st.number_input(
                f"{col}",
                value=default,
                step=step,
                help=f"Range in data: {col_min} – {col_max}" if col_min is not None and col_max is not None else "Enter a number"
            )
        else:
            # Non-numeric fallback
            txt = st.text_input(f"{col} (text)")
            inputs[col] = txt  # will be skipped in scorer if not numeric

    submitted = st.form_submit_button("Find matching solutions")

# --------- Matching Logic ----------
if submitted:
    # Score each row
    results = []
    for i, row in df.iterrows():
        matched, considered = score_row(inputs, row, op_map, i, input_cols)
        ratio = (matched / considered) if considered > 0 else 0.0
        results.append({
            "Solution": row[solution_col],
            "Matched criteria": matched,
            "Criteria considered": considered,
            "Match %": round(100 * ratio, 1),
        })

    res_df = pd.DataFrame(results).sort_values(
        by=["Match %", "Matched criteria"], ascending=[False, False]
    ).reset_index(drop=True)

    st.header("✅ Recommended Solutions")
    if (res_df["Match %"] > 0).any():
        st.success("Here are your top matches (sorted by match % and count):")
    else:
        st.warning("No perfect matches based on numeric criteria. Showing closest matches:")

    st.dataframe(res_df)

    top = res_df.iloc[0] if len(res_df) else None
    if top is not None:
        st.subheader("⭐ Top Pick")
        st.write(f"**{top['Solution']}** – Matched **{top['Matched criteria']}** out of **{top['Criteria considered']}** criteria ({top['Match %']}%).")

# --------- Help & Tips ----------
with st.expander("ℹ️ How matching works (simple explanation)"):
    st.markdown(
        """
        - Each row in your sheet represents a candidate **Solution** with thresholds for each input.
        - For every input column:
          - If the cell is **green**, we check: `your_value ≤ cell_value`.
          - If the cell is **red**, we check: `your_value > cell_value`.
          - If the cell has **no color** or values are **not numeric**, we **skip** that check.
        - We count how many criteria matched for each Solution and sort by best match.
        """
    )
