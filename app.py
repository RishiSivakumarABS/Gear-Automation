import math
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Gearbox Design Tool", layout="wide")

DATA_FILE = Path("data/Gearbox Design Guide Data.xlsx")
STAGE_SHEETS = {
    "Stage 2": "SEN - Stage2",
    "Stage 3": "SZN - Stage3",
    "Stage 4": "SDN - Stage4",
}


def clean_column_name(col: str) -> str:
    return str(col).strip()


@st.cache_data
def load_workbook(file_path: Path) -> dict[str, pd.DataFrame]:
    if not file_path.exists():
        return {}

    excel_file = pd.ExcelFile(file_path)
    sheets = {}

    for stage_label, sheet_name in STAGE_SHEETS.items():
        if sheet_name in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df.columns = [clean_column_name(c) for c in df.columns]
            df = df.dropna(how="all")
            sheets[stage_label] = df

    return sheets


def normalize_size_value(value):
    if pd.isna(value):
        return None
    try:
        number = float(value)
        if number.is_integer():
            return int(number)
        return number
    except Exception:
        return str(value).strip()


def get_size_options(df: pd.DataFrame) -> list:
    if "Size" not in df.columns:
        return []
    return [normalize_size_value(v) for v in df["Size"].dropna().tolist()]


def find_row_by_size(df: pd.DataFrame, selected_size):
    if "Size" not in df.columns:
        return None

    normalized_series = df["Size"].apply(normalize_size_value)
    matches = df[normalized_series == selected_size]

    if matches.empty:
        return None

    return matches.iloc[0]


def format_value(value):
    if pd.isna(value):
        return "-"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.2f}"
    return str(value)


def safe_fourth_root(value: float) -> float:
    return value ** 0.25


def calculate_housing_dimensions(T0: float, Di: float, F: int, a1: float, d2s_ratio: float) -> list[dict]:
    w1_raw = 3.56 * safe_fourth_root(T0)
    w1 = max(w1_raw, 6)
    w2_raw = 0.9 * w1
    w2 = max(w2_raw, 6)
    w3 = 1.5 * w2
    r1 = w1
    r2 = w2
    sR = 2
    phi_i = 1.25 * Di + 10
    sL = 6
    x1 = 2 * w1
    x2 = 2 * w2
    y1 = 3 * w1
    y2 = 3 * w2
    delta_w = 0.6 * w1
    delta_g = 0.4 * w1
    H = 1.06 * a1
    h1 = 2.5 * w1
    h2 = 2.5 * w2

    d2_raw = 3.76 * safe_fourth_root(T0)
    d2 = max(d2_raw, 10)
    d2s = d2s_ratio * d2
    d1_raw = 4.47 * safe_fourth_root(T0) * math.sqrt(2 / F)
    d1 = max(d1_raw, 10)
    d3 = max(0.5 * d2, 8)
    d4 = max(0.5 * d2, 8)
    f1 = 1.5 * d2
    f2 = 1.3 * d2
    p1 = 1.5 * d1
    b2 = 3 * d2
    b1 = 4 * d1

    return [
        {"Name": "Wall thickness of housing base", "Symbol": "w1", "Value": w1, "Unit": "mm", "Remark": ">= 6"},
        {"Name": "Wall thickness of housing cover", "Symbol": "w2", "Value": w2, "Unit": "mm", "Remark": ">= 6"},
        {"Name": "Wall thickness around inspection hole", "Symbol": "w3", "Value": w3, "Unit": "mm", "Remark": ""},
        {"Name": "Base rib thickness next to wall", "Symbol": "r1", "Value": r1, "Unit": "mm", "Remark": ""},
        {"Name": "Cover rib thickness next to wall", "Symbol": "r2", "Value": r2, "Unit": "mm", "Remark": ""},
        {"Name": "Casting slope of ribs", "Symbol": "sR", "Value": sR, "Unit": "deg", "Remark": "approx 1:30"},
        {"Name": "Outer diameter of each bearing lug", "Symbol": "Phi_i", "Value": phi_i, "Unit": "mm", "Remark": "i = 0,1,2,...,n"},
        {"Name": "Casting slope of bearing lugs", "Symbol": "sL", "Value": sL, "Unit": "deg", "Remark": "approx 1:10"},
        {"Name": "Base wall-lug axial transition", "Symbol": "x1", "Value": x1, "Unit": "mm", "Remark": ""},
        {"Name": "Cover wall-lug axial transition", "Symbol": "x2", "Value": x2, "Unit": "mm", "Remark": ""},
        {"Name": "Base wall-lug vertical transition", "Symbol": "y1", "Value": y1, "Unit": "mm", "Remark": ""},
        {"Name": "Cover wall-lug vertical transition", "Symbol": "y2", "Value": y2, "Unit": "mm", "Remark": ""},
        {"Name": "Distance between gear and housing wall", "Symbol": "delta_w", "Value": delta_w, "Unit": "mm", "Remark": ""},
        {"Name": "Distance between gears", "Symbol": "delta_g", "Value": delta_g, "Unit": "mm", "Remark": ""},
        {"Name": "Housing shaft height", "Symbol": "H", "Value": H, "Unit": "mm", "Remark": ""},
        {"Name": "Thickness of housing base lifting hooks", "Symbol": "h1", "Value": h1, "Unit": "mm", "Remark": ""},
        {"Name": "Thickness of housing cover lifting hooks", "Symbol": "h2", "Value": h2, "Unit": "mm", "Remark": "Alternative variant to eye bolts"},
        {"Name": "Diameter of long tie bolts", "Symbol": "d2", "Value": d2, "Unit": "mm", "Remark": ">= 10"},
        {"Name": "Diameter of short tie bolts", "Symbol": "d2s", "Value": d2s, "Unit": "mm", "Remark": f"{d2s_ratio:.2f} x d2"},
        {"Name": "Diameter of foundation bolts", "Symbol": "d1", "Value": d1, "Unit": "mm", "Remark": ">= 10"},
        {"Name": "Diameter of inspection hole lid screws", "Symbol": "d3", "Value": d3, "Unit": "mm", "Remark": ">= 8"},
        {"Name": "Diameter of bearing end cap screws", "Symbol": "d4", "Value": d4, "Unit": "mm", "Remark": ">= 8"},
        {"Name": "Thickness of housing base flange", "Symbol": "f1", "Value": f1, "Unit": "mm", "Remark": ""},
        {"Name": "Thickness of housing cover flange", "Symbol": "f2", "Value": f2, "Unit": "mm", "Remark": ""},
        {"Name": "Thickness of foundation paws", "Symbol": "p1", "Value": p1, "Unit": "mm", "Remark": ""},
        {"Name": "Width of housing flange", "Symbol": "b2", "Value": b2, "Unit": "mm", "Remark": "Check space for wrench"},
        {"Name": "Width of foundation paws", "Symbol": "b1", "Value": b1, "Unit": "mm", "Remark": "Check space for wrench"},
    ]


st.title("Gearbox Design Tool")
st.caption("Lookup stage parameters from Excel or calculate housing dimensions from input values.")

lookup_tab, calculator_tab = st.tabs(["Stage Lookup", "Housing Calculator"])

with lookup_tab:
    workbook_data = load_workbook(DATA_FILE)

    if not workbook_data:
        st.warning(
            "No workbook found yet. Put your Excel file at `data/stage_parameters.xlsx` and make sure the sheet names match the ones in the code."
        )
    else:
        left, right = st.columns([1, 2])

        with left:
            selected_stage = st.selectbox("Stage", list(workbook_data.keys()), key="stage_lookup")
            stage_df = workbook_data[selected_stage]

            if "Size" not in stage_df.columns:
                st.error(f"The sheet for {selected_stage} does not contain a 'Size' column.")
            else:
                size_options = get_size_options(stage_df)
                selected_size = st.selectbox("Size", size_options, key="size_lookup")

        with right:
            st.subheader("Result")
            if "Size" in stage_df.columns:
                row = find_row_by_size(stage_df, selected_size)

                if row is None:
                    st.error("No matching row was found for the selected size.")
                else:
                    result_df = pd.DataFrame(
                        {
                            "Parameter": row.index,
                            "Value": [format_value(v) for v in row.values],
                        }
                    )
                    st.dataframe(result_df, use_container_width=True, hide_index=True)

                    st.subheader("Quick View")
                    quick_cols = st.columns(3)
                    items = list(row.items())
                    for i, (key, value) in enumerate(items):
                        quick_cols[i % 3].metric(label=str(key), value=format_value(value))

with calculator_tab:
    st.subheader("Housing Dimension Calculator")

    col1, col2, col3 = st.columns(3)
    with col1:
        T0 = st.number_input("Torque (T0)", min_value=0.01, value=1000.0, step=10.0)
        Di = st.number_input("Diameter (Di)", min_value=0.01, value=100.0, step=1.0)
    with col2:
        F = st.number_input("Number of bolts (F)", min_value=1, value=8, step=1)
        a1 = st.number_input("a1", min_value=0.01, value=100.0, step=1.0)
    with col3:
        d2s_ratio = st.slider("Short tie bolt factor", min_value=0.70, max_value=1.00, value=0.80, step=0.05)

    calculated_rows = calculate_housing_dimensions(T0=T0, Di=Di, F=int(F), a1=a1, d2s_ratio=d2s_ratio)
    calc_df = pd.DataFrame(calculated_rows)
    calc_df["Value"] = calc_df["Value"].apply(format_value)

    st.dataframe(calc_df, use_container_width=True, hide_index=True)

    st.subheader("Key Outputs")
    metrics = {
        "w1": next(item["Value"] for item in calculated_rows if item["Symbol"] == "w1"),
        "w2": next(item["Value"] for item in calculated_rows if item["Symbol"] == "w2"),
        "d2": next(item["Value"] for item in calculated_rows if item["Symbol"] == "d2"),
        "d1": next(item["Value"] for item in calculated_rows if item["Symbol"] == "d1"),
        "H": next(item["Value"] for item in calculated_rows if item["Symbol"] == "H"),
        "b2": next(item["Value"] for item in calculated_rows if item["Symbol"] == "b2"),
    }
    metric_cols = st.columns(3)
    for i, (label, value) in enumerate(metrics.items()):
        metric_cols[i % 3].metric(label=label, value=format_value(value))

st.divider()
st.markdown(
    """
**Notes**
- The calculator applies the minimum limits shown in your design sheet where needed
- We can next add grouped sections, PDF export, and drawing-based outputs
"""
)