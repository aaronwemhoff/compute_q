# app.py
# ------------------------------
# A beginner-friendly Streamlit app that helps researchers choose a platform
# based on a questionnaire defined in an Excel file.
#
# How it works (high level):
# 1) We read your Excel (headers on Row 4; data on Rows 5–25).
# 2) Each row is a "platform" (e.g., a tool/vendor). The first column is the platform name.
# 3) Each column after the first is a "metric" (e.g., "Max Participants", "Budget", etc.).
# 4) The color of each data cell tells us which way to compare the user input against that platform's capability:
#       - Green cell = the platform works if the user's answer is LESS THAN OR EQUAL TO the cell's value (<=).
#       - Red cell   = the platform works if the user's answer is GREATER THAN the cell's value (>).
#    If a cell is not colored red or green, we default to "greater than or equal to" (>=) to avoid false negatives.
# 5) The user fills out the questionnaire (one number per metric). We compare their inputs to each platform row
#    using the operator indicated by the cell color. Each satisfied metric adds to the platform's score.
# 6) We rank platforms by total score and display the top matches with a simple explanation.
#
# Notes:
# - This app tries to be robust to different shades of red/green.
# - If your sheet has more/less platforms than rows 5–25, you can change those bounds in the sidebar.
# - All code is heavily commented to help beginners learn the flow.
# ------------------------------

import streamlit as st
import pandas as pd
from utils.parser import load_questionnaire, compare_value, detect_operator_label

st.set_page_config(page_title="Research Platform Picker", page_icon="🔎", layout="wide")

# ---- Sidebar: App options
st.sidebar.title("⚙️ App Settings")

st.sidebar.markdown("""
**Excel assumptions (can be changed below):**
- Headers on row **4**
- Data rows from **5** to **25**
- First column = platform name
- Green cell ⇒ use **≤**; Red cell ⇒ use **>**
""")

default_header_row = 4
default_start_row = 5
default_end_row = 25

header_row = st.sidebar.number_input("Header row (1-indexed)", min_value=1, value=default_header_row, step=1)
start_row = st.sidebar.number_input("Start data row (1-indexed)", min_value=1, value=default_start_row, step=1)
end_row = st.sidebar.number_input("End data row (1-indexed)", min_value=start_row, value=default_end_row, step=1)

st.sidebar.divider()
st.sidebar.markdown("### 🧮 Scoring")
st.sidebar.markdown("Each metric that matches the cell's rule contributes 1 point to a platform's score.")
show_weights = st.sidebar.checkbox("Enable metric weights (advanced)", value=False, help="Assign importance to specific metrics.")

st.sidebar.divider()
st.sidebar.markdown("### 📄 Data Source")
uploaded = st.sidebar.file_uploader(
    "Upload your questionnaire Excel (.xlsx)", 
    type=["xlsx"],
    help="If empty, the app will try to load the sample file shipped alongside this demo."
)

# Built-in default path (from this project export). You can replace with your path.
DEFAULT_PATH = "Research Questionnaire Info.xlsx"

# ---- Main Title
st.title("🔎 Research Platform Picker")
st.caption("Answer a few questions and get ranked platform recommendations based on your Excel rules.")

# ---- Load data (either user upload or default file)
xlsx_path = None
if uploaded is not None:
    xlsx_path = uploaded
else:
    # Fall back to default if present in working dir
    xlsx_path = DEFAULT_PATH

try:
    data = load_questionnaire(
        xlsx_path, 
        header_row=header_row, 
        start_row=start_row, 
        end_row=end_row
    )
except Exception as e:
    st.error(f"Could not read the Excel file. Please upload the .xlsx used in your project. Error: {e}")
    st.stop()

platforms_df = data['platforms_df']     # platforms as rows
metrics = data['metrics']               # list of metric names (columns excluding platform name)
ops_grid = data['ops_grid']             # same shape as metrics x platforms, with 'le'/'gt'/'ge' etc
values_grid = data['values_grid']       # numeric values per metric x platform
platform_names = platforms_df.index.tolist()

# ---- Display a preview of parsed data (collapsible)
with st.expander("🔍 Preview parsed data (from Excel)", expanded=False):
    st.write("**Platforms** (rows) and **metrics** (columns). The operators were detected from cell colors.")
    st.dataframe(platforms_df)
    st.write("**Detected comparison operators** (per platform × metric).")
    st.dataframe(pd.DataFrame(ops_grid, index=platform_names, columns=metrics))

# ---- Build questionnaire inputs
st.header("🧾 Questionnaire")

col1, col2 = st.columns([3, 1])
with col1:
    st.write("Provide your requirements for each metric. We'll compare them to each platform's capabilities using the rules from your sheet.")

with col2:
    st.info("Tip: Hover over metric labels to see which operator (≤, >, or ≥) most platforms use.")

# For each metric, compute a reasonable input range from the data (min..max), default to median
metric_inputs = {}
weights = {}

for m in metrics:
    col_vals = platforms_df[m].dropna().astype(float)
    if len(col_vals) == 0:
        min_v, max_v, default_v = 0.0, 100.0, 50.0
    else:
        min_v = float(col_vals.min())
        max_v = float(col_vals.max())
        default_v = float(col_vals.median())
        # Widen the range a little for flexibility
        pad = max(1.0, (max_v - min_v) * 0.1)
        min_v = max(0.0, min_v - pad)
        max_v = max_v + pad

    # Show a hint for the most common operator for this metric across platforms
    # Count which operator shows up most for this column
    op_counts = {}
    for p in platform_names:
        op = ops_grid[p][m]
        op_counts[op] = op_counts.get(op, 0) + 1
    common_op = max(op_counts, key=op_counts.get)

    # UI: number input for the metric value
    metric_inputs[m] = st.number_input(
        f"{m}",
        min_value=float(min_v),
        max_value=float(max_v),
        value=float(default_v),
        step=1.0,
        help=f"Common rule for this metric: {detect_operator_label(common_op)}"
    )

    # Optional weights
    if show_weights:
        weights[m] = st.slider(
            f"Weight for '{m}'",
            min_value=0.0, max_value=5.0, value=1.0, step=0.1,
            help="Set higher weights for more important metrics."
        )
    else:
        weights[m] = 1.0

st.divider()

# ---- Compute scores
st.header("🏆 Recommendations")

results = []
for p in platform_names:
    score = 0.0
    max_score = 0.0
    per_metric_match = {}

    for m in metrics:
        user_val = metric_inputs[m]
        plat_val = values_grid[p][m]
        op = ops_grid[p][m]

        # Full credit if the comparison passes
        ok = compare_value(user_val, plat_val, op)
        w = weights[m]
        max_score += w
        if ok:
            score += w
            per_metric_match[m] = f"✅ match ({detect_operator_label(op).strip()})"
        else:
            per_metric_match[m] = f"❌ no match ({detect_operator_label(op).strip()})"

    pct = 0.0 if max_score == 0 else round(100.0 * score / max_score, 1)
    results.append({
        "Platform": p,
        "Score": round(score, 2),
        "Max Score": round(max_score, 2),
        "Match %": pct,
        "Details": per_metric_match
    })

# Sort by score descending
results_sorted = sorted(results, key=lambda r: (-r["Score"], -r["Match %"], r["Platform"]))

# ---- Show Top N
top_n = st.slider("How many top results to show?", min_value=1, max_value=len(results_sorted), value=min(5, len(results_sorted)))
st.subheader("Top Matches")

for item in results_sorted[:top_n]:
    with st.container(border=True):
        st.markdown(f"### {item['Platform']} — **{item['Match %']}%** match")
        st.progress(item['Match %'] / 100.0)
        st.caption(f"Score: {item['Score']} out of {item['Max Score']} possible")

        with st.expander("See per-metric comparison"):
            # Make a small table for this platform's per-metric status
            det = pd.DataFrame({
                "Metric": list(item["Details"].keys()),
                "Status": list(item["Details"].values()),
                "Your input": [metric_inputs[m] for m in item["Details"].keys()],
                "Platform value": [values_grid[item["Platform"]][m] for m in item["Details"].keys()],
                "Rule": [detect_operator_label(ops_grid[item["Platform"]][m]).strip() for m in item["Details"].keys()],
            })
            st.dataframe(det, hide_index=True, use_container_width=True)

st.divider()
st.markdown("#### Need help?")
st.write("This app is intentionally simple. Check the README for how to adjust parsing rules or scoring if your sheet differs.")

# Footer
st.caption("Built with ❤️ using Streamlit.")

