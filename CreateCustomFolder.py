# app.py
import os
from pathlib import Path
import io
import pandas as pd
import streamlit as st

# ---------- Config ----------
TARGET_WITH_ABBREVIATION = {"Frame Version", "Core Version", "Gasket Version", "Pump"}
SECOND_BLANK_STOP = {"Frame Version", "Core Version", "Gasket Version"}
FIRST_BLANK_STOP = {"Device Name", "Coolant", "Pump"}
TARGET_WITHOUT_ABBREVIATION = {"Device Name", "Coolant"}
SKIP = {"OCCT Version", "OCCT Test Setting", "Fan Settings", "CPU model"}

DEFAULT_FILE = Path("legend_prototype.xlsx")


# ---------- Helpers ----------
def read_legend(excel_path: Path | io.BytesIO, sheet_name=0) -> dict[str, list[str]]:
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except PermissionError:
        # If locked, read from bytes
        with open(excel_path, "rb") as f:
            df = pd.read_excel(io.BytesIO(f.read()), sheet_name=sheet_name)

    df = df.reset_index(drop=True)
    if df.shape[1] < 2:
        raise ValueError("Sheet must have at least two columns (A and B).")

    col_B = df.columns[1]
    results: dict[str, list[str]] = {}
    i, n = 0, len(df)

    while i < n:
        cell = df.at[i, col_B]
        cell_str = str(cell).strip() if pd.notna(cell) else ""

        if cell_str in TARGET_WITH_ABBREVIATION | TARGET_WITHOUT_ABBREVIATION:
            current_cat = cell_str

            # choose column
            if current_cat in TARGET_WITH_ABBREVIATION:
                row_values = df.loc[i].tolist()
                candidates = {"Abbreviation", "Abbrevation"}
                abbrev_index = next(
                    (idx for idx, v in enumerate(row_values) if isinstance(v, str) and v.strip() in candidates),
                    None
                )
                if abbrev_index is None:
                    raise ValueError(f"No 'Abbreviation' column found in row {i}.")
                use_col = df.columns[abbrev_index]
            else:
                use_col = col_B

            vals, blank_seen, j = [], 0, i + 1
            while j < n:
                v = df.at[j, use_col]
                if pd.isna(v) or str(v).strip() == "":
                    blank_seen += 1
                    if (blank_seen == 2 and current_cat in SECOND_BLANK_STOP) or (current_cat in FIRST_BLANK_STOP):
                        break
                else:
                    if v != "?":
                        vals.append(str(v).strip())
                j += 1

            results[current_cat] = vals
            i = j
            continue

        elif cell_str in SKIP:
            i += 1
            continue

        i += 1

    return results


def create_folder(base_dir: str, selections: list[str], time_str: str, server_str: str) -> str:
    folder_name = "_".join(selections + [time_str.strip(), server_str.strip()])
    full_path = os.path.join(base_dir, folder_name)
    os.makedirs(full_path, exist_ok=True)
    return full_path


# ---------- Streamlit UI ----------
st.set_page_config(page_title="Folder Generator", page_icon="📁", layout="centered")
st.title("📁 Folder Generator from Legend Sheet")

# File input
use_default = st.checkbox("Use default legend_prototype.xlsx path", value=True)
uploaded_file = None

if use_default:
    excel_source = DEFAULT_FILE
else:
    uploaded_file = st.file_uploader("Upload legend Excel", type=["xlsx", "xls"])
    if uploaded_file:
        excel_source = io.BytesIO(uploaded_file.read())
    else:
        excel_source = None

# Pick output folder
default_desktop = os.path.join(os.path.expanduser("~"), "Desktop")
out_dir = st.text_input("Base directory to create the new folder in", default_desktop)

# Read legend & build form
OPTIONS = {}
if excel_source:
    try:
        OPTIONS = read_legend(excel_source)
    except Exception as e:
        st.error(f"Failed to read legend: {e}")

if OPTIONS:
    st.subheader("Select configuration options")

    # Keep track of widgets
    selections = {}
    cols = st.columns(1)  # single column, but easy to expand
    for label in OPTIONS:
        with cols[0]:
            selections[label] = st.selectbox(label, [""] + OPTIONS[label], key=label)

    # Additional inputs
    st.markdown("---")
    st.subheader("Add Time and Server Information")

    time_input = st.text_input("Enter time (e.g., 2025-07-23_14-00)")
    server_input = st.text_input("Enter server name (e.g., Server01)")

    # Button
    if st.button("Create Folder"):
        if "" in selections.values() or not time_input.strip() or not server_input.strip():
            st.warning("Please complete all selections, time, and server fields.")
        else:
            try:
                path_created = create_folder(out_dir, list(selections.values()), time_input, server_input)
                st.success(f"Folder created:\n{path_created}")
            except Exception as e:
                st.error(f"Error creating folder: {e}")
else:
    st.info("Load an Excel file to start.")

# --- How to run ---
st.caption("Run with: `streamlit run app.py`")