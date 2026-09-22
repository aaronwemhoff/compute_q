"""
Streamlit app: Research Platform Chooser (Beginner-Friendly)
------------------------------------------------------------
This minimal app helps a researcher answer: Which platform should I use?
Decision rule (for now):
- If information sensitivity is ITAR or PHI -> output "See UTS"
- Otherwise -> output "Use Augie"

All other inputs are captured now for future expansion (e.g., cost, queue fit,
GPU needs, license requirements), but they do not change the decision yet.

How to run locally (once you have Python + pip installed):
1) pip install -r requirements.txt
2) streamlit run app.py

Tip: This file is heavily commented so you can learn by reading top-to-bottom.
"""

# ----------------------
# 1) Imports
# ----------------------
import streamlit as st
from typing import Literal

# ----------------------
# 2) Page config
# ----------------------
st.set_page_config(
    page_title="Research Platform Chooser",
    page_icon="🧭",
    layout="centered",
)

# ----------------------
# 3) Small, typed helpers (totally optional, but beginner-friendly)
# ----------------------
Sensitivity = Literal["open", "ITAR", "PHI"]
SoftwareStatus = Literal[
    "open source",
    "MATLAB",
    "COMSOL",
    "ANSYS",
    "other commercial software",
]
CoreBracket = Literal[
    "16 or less",
    "17 to 192",
    "above 192",
]
WalltimeBracket = Literal[
    "72 hours or less",
    "between 72 and 240 hours",
    "above 240 hours",
]
COMSOLGroupMemBracket = Literal["Yes", "No", "Not using COMSOL"]
ANSYSGroupMemBracket = Literal["Yes", "No", "Not using ANSYS"]
ABAQUSGroupMemBracket = Literal["Yes", "No", "Not using ABAQUS"]

# on/off radio buttons for RAM requirement and GUI

# ----------------------
# 4) App title + short description
# ----------------------
st.title("🧭 Research Platform Chooser")
st.write(
    "This simple advisor collects a few details about your research job and "
    "suggests a computing platform. Various possible outputs are **See UTS**, "
    "**See CRCF**, **vDesktop**, **NSF ACCESS**, **Possible OSG - see CRCF**, "
    "**NSF ACCESS - JetStream2**, **Augie batch w/ long qos**, and **Use Augie**"
)

# ----------------------
# 5) Sidebar: all user inputs live here
# ----------------------
st.sidebar.header("Inputs")

# (1) Information sensitivity level
sensitivity: Sensitivity = st.sidebar.selectbox(
    "Information sensitivity level",
    options=["open", "ITAR", "PHI"],
    help=(
        "Choose how sensitive your information is. ITAR = International Traffic in Arms Regulations; "
        "PHI = Protected Health Information."
    ),
)

# (2) Software status
software_status: SoftwareStatus = st.sidebar.selectbox(
    "Software application",
    options=[
        "open source",
        "MATLAB",
        "COMSOL",
        "ANSYS",
	"ABAQUS",
        "other commercial software",
    ],
    help=(
        "Pick the primary software you plan to use. This doesn't affect the decision yet, "
        "but it's captured for future logic."
    ),
)

# (3a) COMSOL group member?
COMSOLgroupmem: COMSOLGroupMemBracket = st.sidebar.selectbox(
    "Are you a member of the COMSOL user group?",
    options=["No", "Yes"],
)

# (3b) ANSYS group member?
ANSYSgroupmem: ANSYSGroupMemBracket = st.sidebar.selectbox(
    "Are you a member of the ANSYS user group?",
    options=["No", "Yes"],
)

# (3c) ABAQUS group member?
ABAQUSgroupmem: ABAQUSGroupMemBracket = st.sidebar.selectbox(
    "Are you a member of the ABAQUS user group?",
    options=["No", "Yes"],
)

# (4) RAM > 512 GB? (Yes/No)
ram_over_512: bool = st.sidebar.radio(
    "Does your job require more than 512 GB of RAM?",
    options=["No", "Yes"],
    index=0,
    help="If unsure, select No.",
) == "Yes"

# (5) Number of cores bracket
cores: CoreBracket = st.sidebar.selectbox(
    "Number of CPU cores needed",
    options=["16 or less", "17 to 192", "above 192"],
)

# (6) Walltime bracket
walltime: WalltimeBracket = st.sidebar.selectbox(
    "Walltime (how long the job will run)",
    options=["72 hours or less", "between 72 and 240 hours", "above 240 hours"],
)

# (7) Need a GUI (Yes/No)
gui: bool = st.sidebar.radio(
    "Does your job require a Graphical User Interface?",
    options=["No", "Yes"],
    index=0,
    help="Refers only to running on Augie. If unsure, select No.",
) == "Yes"

# ----------------------
# 6) Decision logic (the "rule engine")
# ----------------------
def recommend_platform(info_sensitivity: Sensitivity, cores: CoreBracket, software_status: SoftwareStatus, ram_over_512: bool, walltime: WalltimeBracket, gui: bool, COMSOLgroupmem: COMSOLGroupMemBracket, ANSYSgroupmem: ANSYSGroupMemBracket, ABAQUSgroupmem: ABAQUSGroupMemBracket) -> str:
    """Return the platform recommendation string based on sensitivity and cores."""

    if info_sensitivity in ("ITAR", "PHI"):
        return "See UTS"
    
    if software_status == "other commercial software":
        return "See CRCF"
    elif software_status == "MATLAB":
        if ram_over_512:
            return "Possible NSF ACCESS - see CRCF"
        elif cores == "above 192":
            return "Possible NSF ACCESS - see CRCF"
        elif cores == "16 or less":
            return "Possible OSG - see CRCF"
        elif walltime == "above 240 hours":
            return "Possible NSF ACCESS - see CRCF"
        elif gui:
            return "Possible NSF ACCESS (JetStream2) - see CRCF"
        elif walltime == "between 72 and 240 hours":
            return "Augie batch w/ long qos"
        else:
            return "Augie batch"
    elif software_status == "COMSOL":
        if COMSOLgroupmem == "Yes":
            if ram_over_512:
                return "Possible NSF ACCESS - see CRCF"
            elif cores == "above 192":
                return "Possible NSF ACCESS - see CRCF"
            elif cores == "16 or less":
                return "Possible OSG - see CRCF"
            elif walltime == "above 240 hours":
                return "Possible NSF ACCESS - see CRCF"
            elif gui:
                return "Possible NSF ACCESS (JetStream2) - see CRCF"
            elif walltime == "between 72 and 240 hours":
                return "Augie batch w/ long qos"
            else:
                return "Augie batch"
        else:
            return "See CRCF"
    elif software_status == "ANSYS":
        if ANSYSgroupmem == "Yes":
            if ram_over_512:
                return "Possible NSF ACCESS - see CRCF"
            elif cores == "above 192":
                return "Possible NSF ACCESS - see CRCF"
            elif cores == "16 or less":
                return "Possible OSG - see CRCF"
            elif walltime == "above 240 hours":
                return "Possible NSF ACCESS - see CRCF"
            elif gui:
                return "Possible NSF ACCESS (JetStream2) - see CRCF"
            elif walltime == "between 72 and 240 hours":
                return "Augie batch w/ long qos"
            else:
                return "Augie batch"
        else:
            return "See CRCF"
    elif software_status == "ABAQUS":
        if ABAQUSgroupmem == "Yes":
            if ram_over_512:
                return "Possible NSF ACCESS - see CRCF"
            elif cores == "above 192":
                return "Possible NSF ACCESS - see CRCF"
            elif cores == "16 or less":
                return "Possible OSG - see CRCF"
            elif walltime == "above 240 hours":
                return "Possible NSF ACCESS - see CRCF"
            elif gui:
                return "Possible NSF ACCESS (JetStream2) - see CRCF"
            elif walltime == "between 72 and 240 hours":
                return "Augie batch w/ long qos"
            else:
                return "Augie batch"
        else:
            return "See CRCF"
    else:
        if ram_over_512:
            return "NSF ACCESS"
        elif cores == "above 192":
            return "NSF ACCESS"
        elif cores == "16 or less":
            return "Possible OSG - see CRCF"
        elif walltime == "above 240 hours":
            return "NSF ACCESS"
        elif gui:
            return "NSF ACCESS - JetStream2"
        elif walltime == "between 72 and 240 hours":
            return "Augie batch w/ long qos"
        else:
            return "Augie batch"

# ----------------------
# 7) Main panel: show a summary + recommendation
# ----------------------
with st.expander("What is this doing? (click to expand)"):
    st.markdown(
        """
        **How it works (v1):** The logic behind this app is based
        on a decision flowchart developed by CRCF in September 2025.
        This app implements the flowchart by having users input various
        aspects of their project.
        """
    )

st.subheader("Your inputs")
left, right = st.columns(2)
with left:
    st.write("**Information sensitivity:**", sensitivity)
    st.write("**Software status:**", software_status)
    st.write("**RAM > 512 GB:**", "Yes" if ram_over_512 else "No")
    st.write("**GUI required:**", "Yes" if gui else "NO")
with right:
    st.write("**Cores required:**", cores)
    st.write("**Walltime required:**", walltime)
    st.write("**COMSOL group member:**", COMSOLgroupmem)
    st.write("**ANSYS group member:**", ANSYSgroupmem)
    st.write("**ABAQUS group member:**", ABAQUSgroupmem)

# Button to compute the recommendation (purely for teaching/UX; we could also do it live)
if st.button("Get recommendation"):

    decision = recommend_platform(sensitivity, cores, software_status, ram_over_512, walltime, gui, COMSOLgroupmem, ANSYSgroupmem, ABAQUSgroupmem)

    # Friendly, visual output with an explanation
    if decision == "See UTS":
      st.error("**See UTS**")
      st.caption("Because your information sensitivity is ITAR or PHI, you must use the UTS environment.")
    elif decision == "See CRCF":
      st.info("**See CRCF**")
      st.caption("Your request is unique but may be able to be addressed. Please contact CRCF to discuss.")
    elif decision == "Augie batch w/ long qos":
      st.info("**Augie batch /w long qos**")
      st.caption("Use Augie batch queue with long qos.")
    elif decision == "NSF ACCESS":
      st.info("**NSF ACCESS**")
      st.caption("See CRCF to apply for an NSF ACCESS allocation.")
    elif decision == "NSF ACCESS - JetStream2":
      st.info("**NSF ACCESS - JetStream2")
      st.caption("See CRCF to apply for time on JetStream2.")
    elif decision == "Possible OSG - see CRCF":
      st.info("**Possible OSG - see CRCF")
      st.caption("Your needs may be well met using the Open Science Grid. Please contact CRCF to get more information.")
    elif decision == "Possible NSF ACCESS - see CRCF":
      st.info("**Possible NSF ACCESS - see CRCF")
      st.caption("Your needs may be well met using NSF ACCESS. Please contact CRCF to get more information.")
    elif decision == "Possible NSF ACCESS (JetStream2) - see CRCF":
      st.info("**Possible NSF ACCESS (JetStream2) - see CRCF")
      st.caption("Your needs may be well met using JetStream2. Please contact CRCF to get more information.")
    else:
      st.success("**Use Augie**")
      st.caption("Use the standard Augie batch queue.")

    # (Optional) Show the captured inputs as a dictionary for debugging/learning
    st.code(
        {
            "sensitivity": sensitivity,
            "software_status": software_status,
            "ram_over_512": ram_over_512,
            "cores": cores,
            "walltime": walltime,
        },
        language="python",
    )

# ----------------------
# 8) Footer: tiny stack + next steps (teaching aids)
# ----------------------
st.markdown("---")
st.markdown(
    """
    **Tech stack used here**  
    - UI framework: Streamlit (pure Python, beginner-friendly)  
    - Language: Python 3.10+  
    - Decision engine: a tiny function (`recommend_platform`)  
    - State: ephemeral (no database yet)  

    **Next ideas**  
    - Encode more rules (e.g., RAM/cores/walltime licensing constraints).  
    - Persist past runs to a CSV/SQLite database.  
    - Add role-based guidance and help links.  
    - Deploy to Streamlit Community Cloud or a container.
    """
)
