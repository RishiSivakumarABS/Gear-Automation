import pandas as pd
import streamlit as st
from pathlib import Path

st.set_page_config(page_title="Stage Parameter Lookup", layout="wide")

DATA_FILE = Path("data/Gearbox Design Guide Data.xlsx")
STAGE_SHEETS = {
    "Stage 2": "SEN - Stage2",
    "Stage 3": "SZN - Stage3",
    "Stage 4": "SDN - Stage4",
}


def clean_column_name(col: str) -> str:
    col = str(col).strip()
    return col


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
    sizes = [normalize_size_value(v) for v in df["Size"].dropna().tolist()]
    return sizes



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
    return str(value)


st.title("Stage Parameter Lookup")
st.caption("Select a stage and size to retrieve all related parameters.")

workbook_data = load_workbook(DATA_FILE)

if not workbook_data:
    st.warning(
        "No workbook found yet. Put your Excel file at `data/stage_parameters.xlsx` and make sure the sheet names match the ones in the code."
    )
    st.stop()

left, right = st.columns([1, 2])

with left:
    selected_stage = st.selectbox("Stage", list(workbook_data.keys()))
    stage_df = workbook_data[selected_stage]

    if "Size" not in stage_df.columns:
        st.error(f"The sheet for {selected_stage} does not contain a 'Size' column.")
        st.stop()

    size_options = get_size_options(stage_df)
    selected_size = st.selectbox("Size", size_options)

    lookup_clicked = st.button("Get Parameters", type="primary")

with right:
    st.subheader("Result")

    if lookup_clicked:
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
    else:
        st.info("Choose a stage and size, then click 'Get Parameters'.")

st.divider()
st.subheader("Notes")
st.markdown(
    """
- This first version reads directly from Excel
- Later, we can add search, filtering, PDF export, and calculation tools
"""
)
