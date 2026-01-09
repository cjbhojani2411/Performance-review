import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Performance Review Summary", layout="wide")

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

    df = df.copy()
    df["Month"] = df["Month"].ffill()
    df = df[df["Name"].notna()].copy()

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
    
def detect_header_row(file) -> int:
    """
    Reads the sheet without headers and tries to find the row that contains
    Month/Name/Score (case-insensitive). Returns header row index.
    """
    preview = pd.read_excel(file, header=None, engine="xlrd")
    wanted = {"month", "name", "score"}

    for i in range(min(30, len(preview))):
        row_vals = preview.iloc[i].astype(str).str.strip().str.lower().tolist()
        row_set = set(row_vals)
        if wanted.issubset(row_set):
            return i
    return 1  # fallback (your old default)


st.title("📊 Performance Review System")

header_row = st.sidebar.number_input(
    "Header row (0-based index)", min_value=0, max_value=20, value=1,
    help="Your sheet works with header=1 (second row)."
)

uploaded = st.file_uploader("Upload performance review file (.xls)", type=["xls"])

if not uploaded:
    st.info("Upload an .xls file to generate the summary.")
    st.stop()

# Read XLS
try:
    # df = pd.read_excel(uploaded, header=int(header_row), engine="xlrd")
    # Auto header detection
    auto_header = detect_header_row(uploaded)
    
    # Let user override
    header_row = st.sidebar.number_input(
        "Header row (0-based index)",
        min_value=0, max_value=50, value=int(auto_header),
        help="Auto-detected header row. Change if needed."
    )
    
    # Read again using selected header
    uploaded.seek(0)  # reset file pointer (VERY IMPORTANT)
    df = pd.read_excel(uploaded, header=int(header_row), engine="xlrd")
    
    # Clean column names
    df.columns = df.columns.astype(str).str.strip()

except Exception as e:
    st.error(f"Failed to read XLS: {e}")
    st.stop()

st.subheader("🔍 Raw Sheet Preview")
st.dataframe(df.head(30), width="stretch")

# Generate summary
try:
    summary_df = generate_summary(df)
except Exception as e:
    st.error(str(e))
    st.stop()

st.subheader("✅ Output Sheet")
st.dataframe(summary_df, width="stretch")  # ✅ UI table

# Optional: show in logs + UI text
#with st.expander("See output as text / logs"):
#    st.text(summary_df.to_string(index=False))
#    print("\n=== OUTPUT SHEET ===")
#    print(summary_df.to_string(index=False))

# Download CSV
csv_bytes = summary_df.to_csv(index=False).encode("utf-8")
st.download_button(
    "⬇️ Download CSV",
    data=csv_bytes,
    file_name="performance_monthly_employee_scores.csv",
    mime="text/csv",
)
