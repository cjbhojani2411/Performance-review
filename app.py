import re
import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Performance Review Summary", layout="wide")

# ---------- Helpers ----------
def extract_employee_id(name: str) -> str:
    if pd.isna(name):
        return ""
    m = re.search(r"\bPPS\d+\b", str(name).upper())
    return m.group(0) if m else ""

def clean_employee_name(name: str) -> str:
    if pd.isna(name):
        return ""
    name = str(name)
    name = re.sub(r"\bPPS\d+\b\s*[-–]?\s*", "", name, flags=re.IGNORECASE)
    return name.strip()

def generate_summary(df: pd.DataFrame) -> pd.DataFrame:
    required = {"Month", "Name", "Score"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing columns: {missing}. Found: {list(df.columns)}")

    # Forward-fill Month (your sheet style)
    df = df.copy()
    df["Month"] = df["Month"].ffill()

    # Keep rows that have a Name
    df = df[df["Name"].notna()].copy()

    # Clean score + fields
    df["Score"] = pd.to_numeric(df["Score"], errors="coerce").fillna(0)
    df["EmployeeID"] = df["Name"].apply(extract_employee_id)
    df["Name"] = df["Name"].apply(clean_employee_name)

    summary = (
        df.groupby(["Month", "EmployeeID", "Name"], as_index=False)
          .agg(**{"Average Score": ("Score", "mean")})
    )

    summary["Average Score"] = summary["Average Score"].round(2)
    summary = summary[["Month", "EmployeeID", "Name", "Average Score"]]
    summary = summary.sort_values(["Month", "EmployeeID", "Name"]).reset_index(drop=True)
    return summary

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


# ---------- UI ----------
st.title("📊 Performance Review System (Monthly Average Score)")

st.markdown(
    """
Upload your **performance review .xls** file and generate a clean monthly report:

**Month | EmployeeID | Name | Average Score**
"""
)

with st.sidebar:
    st.header("⚙️ Settings")
    header_row = st.number_input(
        "Header row (0-based index)",
        min_value=0, max_value=20, value=1,
        help="Your file usually works with header=1 (2nd row). Change if columns don't match."
    )

    st.caption("Optional: Save output to disk (server/local machine path)")
    save_to_disk = st.checkbox("Save CSV to output path", value=False)
    output_path = st.text_input(
        "Output CSV path",
        value="/Users/pardypanda/Documents/PPS/Performance review/output/performance_monthly_employee_scores.csv",
        disabled=not save_to_disk
    )

uploaded = st.file_uploader("📁 Upload performance review file (.xls)", type=["xls"])

if not uploaded:
    st.info("Upload an .xls file to continue.")
    st.stop()

# Read Excel
try:
    df = pd.read_excel(uploaded, header=int(header_row), engine="xlrd")
except Exception as e:
    st.error(f"Failed to read the XLS file: {e}")
    st.stop()

st.subheader("🔍 Raw Sheet Preview")
st.dataframe(df.head(30), use_container_width=True)

# Generate
st.subheader("✅ Generated Summary")
try:
    summary_df = generate_summary(df)
except Exception as e:
    st.error(str(e))
    st.stop()

col1, col2, col3 = st.columns(3)
col1.metric("Rows (Raw)", len(df))
col2.metric("Rows (Summary)", len(summary_df))
col3.metric("Employees (Unique)", summary_df["EmployeeID"].nunique())

st.dataframe(summary_df, use_container_width=True)

# Download
csv_bytes = to_csv_bytes(summary_df)

st.download_button(
    label="⬇️ Download CSV",
    data=csv_bytes,
    file_name="performance_monthly_employee_scores.csv",
    mime="text/csv",
)

# Optional save to disk
if save_to_disk:
    try:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        summary_df.to_csv(output_path, index=False)
        st.success(f"✅ Saved to: {output_path}")
    except Exception as e:
        st.warning(f"Could not save file to disk: {e}")

# Small extras
with st.expander("➕ Optional filters"):
    months = sorted(summary_df["Month"].dropna().unique().tolist())
    selected_month = st.selectbox("Filter by Month", ["All"] + months)
    if selected_month != "All":
        st.dataframe(summary_df[summary_df["Month"] == selected_month], use_container_width=True)

    emp_search = st.text_input("Search Name / EmployeeID")
    if emp_search.strip():
        q = emp_search.strip().lower()
        filtered = summary_df[
            summary_df["Name"].str.lower().str.contains(q, na=False) |
            summary_df["EmployeeID"].str.lower().str.contains(q, na=False)
        ]
        st.dataframe(filtered, use_container_width=True)
