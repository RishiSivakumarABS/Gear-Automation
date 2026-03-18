import math
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Gearbox Design Tool", layout="wide")

GUIDE_FILE = Path("data/Gearbox Design Guide Data.xlsx")
POWER_FILE = Path("data/Power to Torque.xlsx")
REFERENCE_IMAGE = Path("assets/gearbox_housing_reference.png")

STAGE_SHEETS = {
    "Stage 1": "SEN - Stage1",
    "Stage 2": "SZN - Stage2",
    "Stage 3": "SDN - Stage3",
}

POWER_STAGE_SHEETS = {
    "Stage 1": "SEN - Stage 1",
    "Stage 2": "SZN - Stage 2",
    "Stage 3": "SDN - Stage 3",
}


def clean_column_name(col: str) -> str:
    return str(col).strip()


@st.cache_data
def load_workbook(file_path: Path, stage_sheet_map: dict[str, str]) -> dict[str, pd.DataFrame]:
    if not file_path.exists():
        return {}

    excel_file = pd.ExcelFile(file_path)
    sheets = {}

    for stage_label, sheet_name in stage_sheet_map.items():
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


def normalize_ratio_value(value):
    if pd.isna(value):
        return None
    try:
        return float(value)
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


def calculate_housing_dimensions(
    T0: float,
    Di: float,
    F: int,
    a1: float,
    d2s_ratio: float,
    B_ratio: float,
    delta_b_ratio: float,
    E: float,
) -> list[dict]:
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

    B = B_ratio * d2
    delta_b = delta_b_ratio * d2

    return [
        {"Name": "Wall thickness of housing base", "Symbol": "w1", "Value": w1, "Unit": "mm", "Remark": ">= 6"},
        {"Name": "Wall thickness of housing cover", "Symbol": "w2", "Value": w2, "Unit": "mm", "Remark": ">= 6"},
        {"Name": "Wall thickness around inspection hole", "Symbol": "w3", "Value": w3, "Unit": "mm", "Remark": ""},
        {"Name": "Base rib thickness next to wall", "Symbol": "r1", "Value": r1, "Unit": "mm", "Remark": ""},
        {"Name": "Cover rib thickness next to wall", "Symbol": "r2", "Value": r2, "Unit": "mm", "Remark": ""},
        {"Name": "Casting slope of ribs", "Symbol": "sR", "Value": sR, "Unit": "deg", "Remark": "approx 1:30"},
        {"Name": "Outer diameter of each bearing lug", "Symbol": "Φi", "Value": phi_i, "Unit": "mm", "Remark": "i = 0,1,2,...,n"},
        {"Name": "Casting slope of bearing lugs", "Symbol": "sL", "Value": sL, "Unit": "deg", "Remark": "approx 1:10"},
        {"Name": "Base wall-lug axial transition", "Symbol": "x1", "Value": x1, "Unit": "mm", "Remark": ""},
        {"Name": "Cover wall-lug axial transition", "Symbol": "x2", "Value": x2, "Unit": "mm", "Remark": ""},
        {"Name": "Base wall-lug vertical transition", "Symbol": "y1", "Value": y1, "Unit": "mm", "Remark": ""},
        {"Name": "Cover wall-lug vertical transition", "Symbol": "y2", "Value": y2, "Unit": "mm", "Remark": ""},
        {"Name": "Distance between gear and housing wall", "Symbol": "δw", "Value": delta_w, "Unit": "mm", "Remark": ""},
        {"Name": "Distance between gears", "Symbol": "δg", "Value": delta_g, "Unit": "mm", "Remark": ""},
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
        {"Name": "Width of bearing lugs", "Symbol": "B", "Value": B, "Unit": "mm", "Remark": "To be confirmed with bearing and end cap dimensions"},
        {"Name": "Distance from tie bolt axis to bearing seat bore", "Symbol": "δb", "Value": delta_b, "Unit": "mm", "Remark": "Check tie bolt hole does not intersect with end cap screw holes"},
        {"Name": "Width of housing", "Symbol": "E", "Value": E, "Unit": "mm", "Remark": "The same for all bearing lugs"},
    ]


def get_power_table_options(power_df: pd.DataFrame):
    if power_df.empty:
        return [], [], []

    required_cols = ["Ratio", "Speed1"]
    for col in required_cols:
        if col not in power_df.columns:
            return [], [], []

    ignored_columns = ["Ratio", "Speed1", "Speed2"]
    size_columns = [col for col in power_df.columns if col not in ignored_columns]

    sizes = [normalize_size_value(col) for col in size_columns]
    ratios = sorted(power_df["Ratio"].dropna().apply(normalize_ratio_value).unique().tolist(), key=float)
    speeds = [speed for speed in [1500, 1000, 750] if speed in power_df["Speed1"].dropna().astype(int).unique().tolist()]

    return sizes, ratios, speeds


def find_power_value(power_df: pd.DataFrame, selected_size, selected_ratio, selected_speed1):
    if power_df.empty:
        return None

    size_col = None
    for col in power_df.columns:
        if col in ["Ratio", "Speed1", "Speed2"]:
            continue
        if normalize_size_value(col) == selected_size:
            size_col = col
            break

    if size_col is None:
        return None

    ratio_series = power_df["Ratio"].apply(normalize_ratio_value)
    speed_series = pd.to_numeric(power_df["Speed1"], errors="coerce")

    matches = power_df[
        (ratio_series == selected_ratio) &
        (speed_series == selected_speed1)
    ]

    if matches.empty:
        return None

    power_value = matches.iloc[0][size_col]
    if pd.isna(power_value):
        return None

    return float(power_value)


def calculate_output_torque(power_value: float, speed1: float) -> float:
    return (power_value * 60) / (2 * math.pi * speed1)


st.title("Gearbox Design Tool")
st.caption("Convert power to torque, lookup stage parameters, or calculate housing dimensions.")

power_tab, lookup_tab, calculator_tab = st.tabs(
    ["Power to Torque", "Stage Lookup", "Housing Calculator"]
)

with power_tab:
    st.subheader("Power to Torque")

    power_workbook = load_workbook(POWER_FILE, POWER_STAGE_SHEETS)

    if not power_workbook:
        st.warning(
            "Could not load the Power to Torque workbook. Make sure `data/Power to Torque.xlsx` exists and the sheet names match the code."
        )
    else:
        left, right = st.columns([1, 2])

        with left:
            selected_power_stage = st.selectbox(
                "Stage",
                list(power_workbook.keys()),
                key="power_stage",
            )

            power_df = power_workbook[selected_power_stage]
            size_options, ratio_options, speed_options = get_power_table_options(power_df)

            selected_power_size = st.selectbox("Size", size_options, key="power_size")
            selected_power_ratio = st.selectbox("Ratio", ratio_options, key="power_ratio")

            allowed_speeds = [speed for speed in [1500, 1000, 750] if speed in speed_options]
            selected_speed1 = st.selectbox("Speed1", allowed_speeds, key="power_speed")

        with right:
            st.subheader("Result")

            power_value = find_power_value(
                power_df,
                selected_power_size,
                selected_power_ratio,
                selected_speed1,
            )

            if power_value is None:
                st.error("No matching power value was found for the selected Stage, Size, Ratio, and Speed1.")
            else:
                output_torque = calculate_output_torque(power_value, selected_speed1)

                result_data = pd.DataFrame(
                    [
                        {"Parameter": "Stage", "Value": selected_power_stage},
                        {"Parameter": "Size", "Value": format_value(selected_power_size)},
                        {"Parameter": "Ratio", "Value": format_value(selected_power_ratio)},
                        {"Parameter": "Speed1", "Value": format_value(selected_speed1)},
                        {"Parameter": "Power", "Value": format_value(power_value)},
                        {"Parameter": "Output Torque", "Value": format_value(output_torque)},
                    ]
                )

                st.dataframe(result_data, use_container_width=True, hide_index=True)

                metric_cols = st.columns(2)
                metric_cols[0].metric("Power", format_value(power_value))
                metric_cols[1].metric("Output Torque", format_value(output_torque))

                st.markdown("### Formula Used")
                st.latex(r"T = \frac{P \cdot 60}{2\pi \cdot n_1}")

with lookup_tab:
    workbook_data = load_workbook(GUIDE_FILE, STAGE_SHEETS)

    if not workbook_data:
        st.warning(
            "No workbook found yet. Put your Excel file at `data/Gearbox Design Guide Data.xlsx` and make sure the sheet names match the ones in the code."
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

    input_col, image_col = st.columns([1, 1])

    with input_col:
        st.markdown("### Inputs")

        col1, col2, col3 = st.columns(3)
        with col1:
            T0 = st.number_input("Torque (T0)", min_value=0.01, value=1000.0, step=10.0)
            Di = st.number_input("Diameter (Di)", min_value=0.01, value=100.0, step=1.0)
            E = st.number_input("Width of housing (E)", min_value=0.01, value=200.0, step=1.0)

        with col2:
            F = st.number_input("Number of bolts (F)", min_value=1, value=8, step=1)
            a1 = st.number_input("a1", min_value=0.01, value=100.0, step=1.0)

        with col3:
            d2s_ratio = st.slider(
                "Short tie bolt ratio (d2s / d2)",
                min_value=0.70,
                max_value=1.00,
                value=0.80,
                step=0.05,
            )
            B_ratio = st.slider(
                "Bearing lug width ratio (B / d2)",
                min_value=3.0,
                max_value=3.5,
                value=3.2,
                step=0.1,
            )
            delta_b_ratio = st.slider(
                "Bolt axis distance ratio (δb / d2)",
                min_value=1.0,
                max_value=1.2,
                value=1.1,
                step=0.05,
            )

    with image_col:
        st.markdown("### Reference Drawing")
        if REFERENCE_IMAGE.exists():
            st.image(
                str(REFERENCE_IMAGE),
                caption="Gearbox housing dimension reference",
                use_container_width=True,
            )
        else:
            st.info("Add the reference drawing at `assets/gearbox_housing_reference.png` to display it here.")

    calculated_rows = calculate_housing_dimensions(
        T0=T0,
        Di=Di,
        F=int(F),
        a1=a1,
        d2s_ratio=d2s_ratio,
        B_ratio=B_ratio,
        delta_b_ratio=delta_b_ratio,
        E=E,
    )

    calc_df = pd.DataFrame(calculated_rows)
    calc_df["Value"] = calc_df["Value"].apply(format_value)

    st.subheader("Calculated Results")
    st.dataframe(calc_df, use_container_width=True, hide_index=True)

    st.subheader("Key Outputs")
    metrics = {
        "w1": next(item["Value"] for item in calculated_rows if item["Symbol"] == "w1"),
        "w2": next(item["Value"] for item in calculated_rows if item["Symbol"] == "w2"),
        "d2": next(item["Value"] for item in calculated_rows if item["Symbol"] == "d2"),
        "d1": next(item["Value"] for item in calculated_rows if item["Symbol"] == "d1"),
        "B": next(item["Value"] for item in calculated_rows if item["Symbol"] == "B"),
        "δb": next(item["Value"] for item in calculated_rows if item["Symbol"] == "δb"),
    }

    metric_cols = st.columns(3)
    for i, (label, value) in enumerate(metrics.items()):
        metric_cols[i % 3].metric(label=label, value=format_value(value))

st.divider()
st.markdown(
    """
**Notes**
- The Power to Torque workbook is loaded from `data/Power to Torque.xlsx`
- Power to Torque uses Stage, Size, Ratio, and Speed1 as inputs
- Stage lookup uses `data/Gearbox Design Guide Data.xlsx`
- The housing calculator applies the minimum limits shown in the design sheet where needed
- `B`, `δb`, and `d2s` are calculated as a factor times `d2`
- `E` is treated as a manual design input
"""
)