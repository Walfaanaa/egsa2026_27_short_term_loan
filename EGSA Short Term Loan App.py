# EGSA Short Term Loan App - CSV/Excel Compatible

import streamlit as st
import pandas as pd
from datetime import date
from pathlib import Path
from io import BytesIO


# ==========================
# Page Configuration
# ==========================
st.set_page_config(
    page_title="EGSA Short Term Loan App",
    layout="wide"
)

st.title("EGSA Short Term Loan Management")


# ==========================
# Upload File (Optional)
# ==========================
uploaded_file = st.file_uploader(
    "Upload your Excel or CSV file (Optional)",
    type=["xlsx", "xls", "csv"]
)


# ==========================
# Load Data
# ==========================
if uploaded_file is not None:

    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)

    else:
        df = pd.read_excel(uploaded_file)

    st.success("Uploaded file loaded successfully.")

else:

    default_file = Path("EGSA2026_27_short_loan.xlsx")

    if default_file.exists():

        df = pd.read_excel(default_file)

        st.info(
            "Using default file: EGSA2026_27_short_loan.xlsx"
        )

    else:

        st.error(
            "EGSA2026_27_short_loan.xlsx not found."
        )

        st.stop()


# ==========================
# Clean Column Names
# ==========================
df.columns = (
    df.columns
    .str.strip()
    .str.lower()
)


# ==========================
# Required Columns
# ==========================
required_columns = [
    "loan_id",
    "business_date",
    "id",
    "loan_type",
    "disbursed_amount",
    "interest_rate",
    "interest_amount",
    "collection_amount",
    "from_account",
    "to_account",
    "due_date",
    "phone_no",
    "status"
]


# ==========================
# Check Missing Columns
# ==========================
missing_cols = [
    col for col in required_columns
    if col not in df.columns
]


if missing_cols:

    st.error(
        f"Missing columns in your file: {missing_cols}"
    )

    st.write("Available columns:")
    st.write(df.columns.tolist())

    st.stop()



# ==========================
# Validate Loan Type
# Level 1 - Level 6
# ==========================
valid_types = [
    f"level_{i}"
    for i in range(1, 7)
]


df["loan_type"] = (
    df["loan_type"]
    .astype(str)
    .str.lower()
)


df = df[
    df["loan_type"].isin(valid_types)
]



# ==========================
# Default Status
# ==========================
df["status"] = (
    df["status"]
    .fillna("in progress")
    .astype(str)
    .str.lower()
)



# ==========================
# Convert Dates
# ==========================
df["business_date"] = pd.to_datetime(
    df["business_date"],
    errors="coerce"
)


df["due_date"] = pd.to_datetime(
    df["due_date"],
    errors="coerce"
)



# ==========================
# Auto Update Loan Status
# ==========================
today = pd.to_datetime(date.today())


df.loc[
    (df["due_date"] <= today)
    &
    (df["status"] == "in progress"),
    "status"
] = "completed"



# ==========================
# Loan Summary
# ==========================
st.subheader("Loan Summary")


col1, col2, col3 = st.columns(3)


col1.metric(
    "Total Loans",
    len(df)
)


col2.metric(
    "In Progress",
    len(
        df[
            df["status"] == "in progress"
        ]
    )
)


col3.metric(
    "Completed",
    len(
        df[
            df["status"] == "completed"
        ]
    )
)



# ==========================
# Loan Type Summary
# ==========================
st.subheader("Loan Level Summary")


loan_summary = (
    df.groupby("loan_type")
    .agg(
        total_loans=("loan_id", "count"),
        total_disbursed=("disbursed_amount", "sum")
    )
    .reset_index()
)


st.dataframe(
    loan_summary,
    use_container_width=True
)



# ==========================
# Display Loan Details
# ==========================
tab1, tab2 = st.tabs(
    [
        "In Progress Loans",
        "Completed Loans"
    ]
)


with tab1:

    st.dataframe(
        df[
            df["status"] == "in progress"
        ],
        use_container_width=True
    )


with tab2:

    st.dataframe(
        df[
            df["status"] == "completed"
        ],
        use_container_width=True
    )



# ==========================
# Download Updated Excel
# ==========================
def to_excel(data):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        data.to_excel(
            writer,
            index=False,
            sheet_name="Loans"
        )

    return output.getvalue()



st.download_button(
    label="📥 Download Updated Excel",
    data=to_excel(df),
    file_name="EGSA_short_term_loans_updated.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    )
)
