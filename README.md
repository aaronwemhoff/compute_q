# Research Platform Picker (Streamlit + Python)

A beginner-friendly Streamlit app to help researchers select the best platform based on a simple questionnaire defined in an Excel file.

## 🧠 How it works

Your Excel file sets the "rules of the game":

- **Headers** are on **Row 4** (1-indexed). The first header is the platform name; the remaining headers are **metrics** (e.g., Budget, Max Participants, etc.).  
- **Platforms** are the rows from **Row 5** through **Row 25** (inclusive). Each row corresponds to a platform/tool.  
- **Cell color indicates rule direction** for comparing user input against the platform's value for that metric:  
  - **Green** ⇒ the user's input must be **≤** the cell value (`user ≤ platform`).  
  - **Red** ⇒ the user's input must be **>** the cell value (`user > platform`).  
  - Uncolored cells default to **≥** (`user ≥ platform`) as a safe fallback.

The app reads this structure and builds a questionnaire UI automatically. Each satisfied metric contributes **1 point** to a platform's score. You can optionally enable **weights** for metrics in the sidebar.

## ▶️ Quick start

1. **Install Python 3.9+** (3.10/3.11 recommended).
2. Create and activate a virtual environment (optional but recommended).
3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Make sure your Excel file is available in the same directory as `app.py` (default name: `Research Questionnaire Info.xlsx`).  
   Or upload it via the sidebar when the app is running.
5. Start the app:

```bash
streamlit run app.py
```

## 📁 Expected Excel layout

- Row **4** has column names: the first is the platform label, the rest are metric names.
- Rows **5–25** have platform data. You can change these bounds via the sidebar if needed.
- The app attempts to detect **green**/**red** fills across common shades. If your theme uses special tints, you can update the `GREEN_HINTS` / `RED_HINTS` sets in `utils/parser.py`.

## 🧩 Customization ideas

- **Scoring**: Change `compare_value` or add partial-credit logic (e.g., distance to threshold).
- **Operators**: Extend `_operator_from_color` to support other colors meaning different rules (e.g., blue = equality).
- **Visuals**: Add charts of platform scores, export results to CSV, or show per-metric tooltips with source cell references.
- **Validation**: Add min/max bounds, required fields, or custom default values for each metric.

## ❓Troubleshooting

- If the app says it can't find headers or platforms, verify your rows match the defaults (or adjust in the sidebar).
- If colors don't register, inspect a cell's RGB in `utils/parser.py::_read_rgb` and include its hex in the hint sets.
- If values show as `NaN`, ensure the cells contain numbers (not text).

---

Built with ❤️ using Streamlit.
