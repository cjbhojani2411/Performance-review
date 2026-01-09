# performance_summary_csv.py
# pip install pandas xlrd==2.0.1

import os
import re
import pandas as pd

INPUT_XLS = r"/Users/pardypanda/Documents/PPS/Performance review/source/performance review.xls"
OUTPUT_CSV = r"/Users/pardypanda/Documents/PPS/Performance review/output/performance_monthly_employee_scores.csv"


def extract_employee_id(name: str) -> str:
    if pd.isna(name):
        return ""
    m = re.search(r"\bPPS\d+\b", str(name).upper())
    return m.group(0) if m else ""


def clean_employee_name(name: str) -> str:
    """
    Removes employee code like 'PPS015 - ' from Name
    """
    if pd.isna(name):
        return ""
    name = str(name)
    # remove PPSxxx and surrounding separators
    name = re.sub(r"\bPPS\d+\b\s*[-–]?\s*", "", name, flags=re.IGNORECASE)
    return name.strip()


def main():
    if not os.path.exists(INPUT_XLS):
        raise FileNotFoundError(f"File not found: {INPUT_XLS}")

    df = pd.read_excel(INPUT_XLS, header=1, engine="xlrd")

    required = {"Month", "Name", "Score"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing columns {missing}. Found: {list(df.columns)}")

    # Forward-fill Month (important)
    df["Month"] = df["Month"].ffill()

    # Keep valid rows
    df = df[df["Name"].notna()].copy()

    # Clean fields
    df["Score"] = pd.to_numeric(df["Score"], errors="coerce").fillna(0)
    df["EmployeeID"] = df["Name"].apply(extract_employee_id)
    df["Name"] = df["Name"].apply(clean_employee_name)

    summary = (
        df.groupby(["Month", "EmployeeID", "Name"], as_index=False)
          .agg(
              **{
                  "Average Score": ("Score", "mean"),
              }
          )
    )

    summary["Average Score"] = summary["Average Score"].round(2)

    # Final output format (NO Total Score)
    summary = summary[
        ["Month", "EmployeeID", "Name", "Average Score"]
    ].sort_values(["Month", "EmployeeID", "Name"])

    summary.to_csv(OUTPUT_CSV, index=False)
    print("✅ CSV generated:", OUTPUT_CSV)


if __name__ == "__main__":
    main()
