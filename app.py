# app.py
# ------------------------------
# Updated Streamlit app that handles mixed data types (numeric and categorical)
# from Excel questionnaire for research platform selection.
#
# Key improvements:
# 1) Handles both numeric and text data appropriately
# 2) Uses dropdowns for categorical data (like OS, software, data sensitivity)
# 3) Uses number inputs for numeric data (like cores, memory, walltime)
# 4) Properly compares user inputs against platform capabilities based on data type
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
- Mixed data types supported (numeric & text)
""")

default_header_row = 4
default_start_row = 5
default_end_row = 25

header_row = st.sidebar.number_input("Header row (1-indexed)", min_value=1, value=default_header_row, step=1)
start_row = st.sidebar.number_input("Start data row (1-indexed)", min_value=1, value=default_start_row, step=1)
end_row = st.sidebar.number_input("End data row (1-indexed)", min_value=start_row, value=default_end_row, step=1)

st.sidebar.divider()
st.sidebar.markdown("### 🧮 Scoring")
st.sidebar.markdown("Each metric that matches the platform's rule contributes 1 point to the platform's score.")
show_weights = st.sidebar.checkbox("Enable metric weights (advanced)", value=False, help="Assign importance to specific metrics.")

st.sidebar.divider()
st.sidebar.markdown("### 📄 Data Source")
uploaded = st.sidebar.file_uploader(
    "Upload your questionnaire Excel (.xlsx)", 
    type=["xlsx"],
    help="Upload your Excel file with platform data and colored cells indicating comparison rules."
)

# Built-in default path
DEFAULT_PATH = "Research Questionnaire Info.xlsx"

# ---- Main Title
st.title("🔎 Research Platform Picker")
st.caption("Answer questions about your research needs and get ranked platform recommendations based on your Excel rules.")

# ---- Load data (either user upload or default file)
xlsx_path = uploaded if uploaded is not None else DEFAULT_PATH

try:
    data = load_questionnaire(
        xlsx_path, 
        header_row=header_row, 
        start_row=start_row, 
        end_row=end_row
    )
except Exception as e:
    st.error(f"Could not read the Excel file. Please upload the correct .xlsx file. Error: {e}")
    st.stop()

platforms_df = data['platforms_df']     # platforms as rows
metrics = data['metrics']               # list of metric names (columns excluding platform name)
ops_grid = data['ops_grid']             # operators per platform x metric
values_grid = data['values_grid']       # values per platform x metric
data_types = data['data_types']         # data type per metric ('numeric' or 'categorical')
platform_names = platforms_df.index.tolist()

# ---- Display a preview of parsed data (collapsible)
with st.expander("🔍 Preview parsed data (from Excel)", expanded=False):
    st.write("**Platforms** (rows) and **metrics** (columns) with detected data types:")
    
    # Show data types
    type_info = pd.DataFrame({
        'Metric': metrics,
        'Data Type': [data_types[m] for m in metrics]
    })
    st.dataframe(type_info, hide_index=True)
    
    st.write("**Platform data:**")
    st.dataframe(platforms_df)
    
    st.write("**Detected comparison operators** (per platform × metric):")
    st.dataframe(pd.DataFrame(ops_grid, index=platform_names, columns=metrics))

# ---- Build questionnaire inputs
st.header("🧾 Questionnaire")

col1, col2 = st.columns([3, 1])
with col1:
    st.write("Provide your requirements for each metric. The app will compare them to each platform's capabilities using the rules from your Excel sheet.")

with col2:
    st.info("💡 **Tip:** Green cells in Excel = '≤' comparison, Red cells = '>' comparison. Categorical data uses exact matching.")

# For each metric, create appropriate input widget based on data type
# Skip "Solution" column as it's the recommendation output, not user input
metric_inputs = {}
weights = {}

for m in metrics:
    # Skip solution column - it's the output, not input
    if 'solution' in m.lower():
        continue
        
    col_vals = platforms_df[m].dropna()
    data_type = data_types[m]
    
    st.subheader(f"📊 {m}")
    
    if data_type == 'categorical':
        # For categorical data, create a selectbox with unique values
        unique_values = sorted([str(v) for v in col_vals.unique() if pd.notna(v)])
        
        # Add an "Other" option
        unique_values.append("Other")
        
        # Set default values for specific metrics
        default_index = 0
        if m.lower() in ['data sensitivity']:
            if "Open" in unique_values:
                default_index = unique_values.index("Open")
        elif m.lower() in ['software', 'software package']:
            if "open-source" in unique_values:
                default_index = unique_values.index("open-source")
        
        selected = st.selectbox(
            f"Select your requirement for {m}:",
            options=unique_values,
            index=default_index,
            key=f"select_{m}"
        )
        
        if selected == "Other":
            metric_inputs[m] = st.text_input(
                f"Enter custom value for {m}:",
                key=f"custom_{m}"
            )
        else:
            metric_inputs[m] = selected
            
        # Show common operator for this metric
        op_counts = {}
        for p in platform_names:
            op = ops_grid[p][m]
            op_counts[op] = op_counts.get(op, 0) + 1
        common_op = max(op_counts, key=op_counts.get) if op_counts else 'eq'
        
        st.caption(f"ℹ️ Most platforms use: {detect_operator_label(common_op)}")
        
    else:
        # For numeric data, create number input with reasonable bounds
        numeric_vals = []
        for val in col_vals:
            try:
                numeric_vals.append(float(val))
            except (ValueError, TypeError):
                pass
        
        if numeric_vals:
            min_v = float(min(numeric_vals))
            max_v = float(max(numeric_vals))
            default_v = float(pd.Series(numeric_vals).median())
            # Widen the range for flexibility
            pad = max(1.0, (max_v - min_v) * 0.2)
            min_v = max(0.0, min_v - pad)
            max_v = max_v + pad
        else:
            min_v, max_v, default_v = 0.0, 100.0, 50.0

        # Set specific default values for certain metrics
        if 'walltime' in m.lower() and 'hr' in m.lower():
            default_v = 24.0
        elif 'runs' in m.lower() and '#' in m.lower():
            default_v = 1.0

        metric_inputs[m] = st.number_input(
            f"Enter your requirement for {m}:",
            min_value=float(min_v),
            max_value=float(max_v),
            value=float(default_v),
            step=1.0 if max_v > 10 else 0.1,
            key=f"number_{m}"
        )
        
        # Show common operator for this metric
        op_counts = {}
        for p in platform_names:
            op = ops_grid[p][m]
            op_counts[op] = op_counts.get(op, 0) + 1
        common_op = max(op_counts, key=op_counts.get) if op_counts else 'ge'
        
        st.caption(f"ℹ️ Most platforms use: {detect_operator_label(common_op)}")

    # Optional weights
    if show_weights:
        weights[m] = st.slider(
            f"Importance weight for '{m}'",
            min_value=0.0, max_value=5.0, value=1.0, step=0.1,
            help="Higher weights make this metric more important in the scoring.",
            key=f"weight_{m}"
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
    recommended_solution = ""

    for m in metrics:
        # Skip solution column in scoring - it's the output
        if 'solution' in m.lower():
            recommended_solution = str(values_grid[p][m]) if pd.notna(values_grid[p][m]) else "N/A"
            continue
            
        user_val = metric_inputs[m]
        plat_val = values_grid[p][m]
        op = ops_grid[p][m]
        data_type = data_types[m]

        # Check if the comparison passes
        ok = compare_value(user_val, plat_val, op, data_type)
        w = weights[m]
        max_score += w
        
        if ok:
            score += w
            per_metric_match[m] = f"✅ Match ({detect_operator_label(op).split('(')[0].strip()})"
        else:
            per_metric_match[m] = f"❌ No match ({detect_operator_label(op).split('(')[0].strip()})"

    pct = 0.0 if max_score == 0 else round(100.0 * score / max_score, 1)
    results.append({
        "Platform": p,
        "Score": round(score, 2),
        "Max Score": round(max_score, 2),
        "Match %": pct,
        "Details": per_metric_match,
        "Recommended Solution": recommended_solution
    })

# Sort by score descending, then by match percentage, then alphabetically
results_sorted = sorted(results, key=lambda r: (-r["Score"], -r["Match %"], r["Platform"]))

# ---- Show Top N
if results_sorted:
    top_n = st.slider("How many top results to show?", min_value=1, max_value=len(results_sorted), value=min(5, len(results_sorted)))
    
    st.subheader(f"🥇 Top {top_n} Recommendations")

    for i, item in enumerate(results_sorted[:top_n], 1):
        # Color code the containers based on match percentage
        if item['Match %'] >= 80:
            border_color = "green"
        elif item['Match %'] >= 60:
            border_color = "orange" 
        else:
            border_color = "red"
            
        with st.container(border=True):
            col1, col2 = st.columns([3, 1])
            
            with col1:
                st.markdown(f"### #{i} {item['Platform']}")
                st.markdown(f"**{item['Match %']}%** compatibility match")
                if item['Recommended Solution']:
                    st.markdown(f"🎯 **Recommended:** {item['Recommended Solution']}")
            
            with col2:
                st.metric("Score", f"{item['Score']}/{item['Max Score']}")
            
            # Progress bar
            progress_color = "normal"
            if item['Match %'] >= 80:
                progress_color = "normal"  # Green
            elif item['Match %'] >= 60:
                progress_color = "normal"  # Will appear as default blue
            
            st.progress(item['Match %'] / 100.0)

            with st.expander("📋 Detailed comparison"):
                # Create detailed comparison table
                comparison_data = []
                for metric in metrics:
                    # Skip solution column in detailed comparison
                    if 'solution' in metric.lower():
                        continue
                        
                    user_input = metric_inputs[metric]
                    platform_value = values_grid[item["Platform"]][metric]
                    rule = detect_operator_label(ops_grid[item["Platform"]][metric]).split('(')[0].strip()
                    status = item["Details"][metric]
                    
                    comparison_data.append({
                        "Metric": metric,
                        "Your Input": str(user_input),
                        "Platform Value": str(platform_value) if pd.notna(platform_value) else "N/A",
                        "Comparison Rule": rule,
                        "Result": status,
                        "Data Type": data_types[metric].title()
                    })
                
                comp_df = pd.DataFrame(comparison_data)
                st.dataframe(comp_df, hide_index=True, use_container_width=True)

else:
    st.warning("No platforms found in the data. Please check your Excel file.")

# ---- Summary statistics
if results_sorted:
    st.divider()
    st.subheader("📊 Summary Statistics")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Platforms", len(results_sorted))
    
    with col2:
        high_match = len([r for r in results_sorted if r['Match %'] >= 80])
        st.metric("High Match (≥80%)", high_match)
    
    with col3:
        med_match = len([r for r in results_sorted if 60 <= r['Match %'] < 80])
        st.metric("Medium Match (60-79%)", med_match)
    
    with col4:
        low_match = len([r for r in results_sorted if r['Match %'] < 60])
        st.metric("Low Match (<60%)", low_match)

st.divider()
st.markdown("### 💡 How it works")
st.markdown("""
- **Green cells** in your Excel file mean the platform works if your input is **≤** the platform's value
- **Red cells** mean the platform works if your input is **>** the platform's value  
- **Uncolored cells** default to **≥** for numeric data or **exact match** for text data
- Each matching criterion adds points to the platform's score
- Results are ranked by total score and match percentage
""")

# Footer
st.caption("Built with ❤️ using Streamlit | Updated to handle mixed data types")
