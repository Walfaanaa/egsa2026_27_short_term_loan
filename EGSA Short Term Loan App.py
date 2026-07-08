# ==================== EGSA SHORT TERM LOAN APP ====================

import streamlit as st
import pandas as pd
from io import BytesIO


# -------------------- CONFIG --------------------
st.set_page_config(
    page_title="EGSA Short Term Loan App",
    layout="wide"
)

st.title("📊 EGSA Short Term Loan Management")


# -------------------- SESSION STATE --------------------
if "data_source" not in st.session_state:
    st.session_state.data_source = None


# -------------------- INPUT --------------------

st.sidebar.header("📂 Data Source")

uploaded_file = st.sidebar.file_uploader(
    "Upload Local Excel or CSV",
    type=["xlsx", "xls", "csv"]
)


github_url = (
    "https://raw.githubusercontent.com/"
    "Walfaanaa/egsa2026_27_short_term_loan/"
    "main/EGSA2026_27_short_loan.xlsx"
)


if st.sidebar.button("🌐 Load From GitHub"):
    st.session_state.data_source = github_url


if uploaded_file is not None:
    st.session_state.data_source = uploaded_file



# -------------------- LOAD DATA --------------------
def load_data(source):

    try:

        if isinstance(source, str):

            if source.endswith(".xlsx"):
                return pd.read_excel(source)

            else:
                return pd.read_csv(source)

        else:

            if source.name.endswith((".xlsx", ".xls")):
                return pd.read_excel(source)

            else:
                return pd.read_csv(source)


    except Exception as e:

        st.error(f"Error reading file: {e}")

        return None



# -------------------- PROCESS --------------------

if st.session_state.data_source:


    df = load_data(
        st.session_state.data_source
    )


    if df is not None:


        # -------------------- CLEAN COLUMN NAME --------------------

        df.columns = (
            df.columns
            .str.strip()
            .str.lower()
        )


        # -------------------- REQUIRED COLUMNS --------------------

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


        missing = [
            c for c in required_columns
            if c not in df.columns
        ]


        if missing:

            st.error(
                f"Missing columns: {missing}"
            )

            st.write(
                "Available columns:",
                df.columns.tolist()
            )

            st.stop()



        # -------------------- LOAN TYPE --------------------

        valid_types = [
            f"level_{i}"
            for i in range(1,7)
        ]


        df["loan_type"] = (
            df["loan_type"]
            .astype(str)
            .str.lower()
        )


        df = df[
            df["loan_type"].isin(valid_types)
        ]



        # -------------------- CLEAN DATA --------------------

        df["status"] = (
            df["status"]
            .fillna("in progress")
            .astype(str)
            .str.lower()
        )


        df["due_date"] = pd.to_datetime(
            df["due_date"],
            errors="coerce"
        )


        for col in [
            "disbursed_amount",
            "interest_amount",
            "collection_amount"
        ]:

            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            ).fillna(0)



        # -------------------- DAYS CALCULATION --------------------

        today = pd.Timestamp.today().normalize()

        df["days_left"] = (
            df["due_date"] - today
        ).dt.days



        # -------------------- STATUS LOGIC --------------------

        df.loc[
            (df["days_left"] <= 0)
            &
            (df["collection_amount"] >= df["disbursed_amount"]),
            "status"
        ] = "completed"


        df.loc[
            (df["days_left"] <= 0)
            &
            (df["collection_amount"] < df["disbursed_amount"]),
            "status"
        ] = "overdue"


        df.loc[
            df["days_left"] == 1,
            "status"
        ] = "1 day left"


        df.loc[
            df["days_left"] == 2,
            "status"
        ] = "2 days left"


        df.loc[
            df["days_left"] > 2,
            "status"
        ] = "in progress"



        # -------------------- FILTER --------------------

        st.sidebar.header("🔍 Filter")

        status_filter = st.sidebar.multiselect(
            "Status",
            df["status"].unique(),
            default=df["status"].unique()
        )


        search_id = st.sidebar.text_input(
            "Search ID"
        )


        filtered_df = df[
            df["status"].isin(status_filter)
        ]


        if search_id:

            filtered_df = filtered_df[
                filtered_df["id"]
                .astype(str)
                .str.contains(search_id)
            ]



        # -------------------- SUMMARY --------------------

        st.subheader("📌 Loan Summary")


        c1,c2,c3,c4,c5 = st.columns(5)


        c1.metric(
            "Total Loans",
            len(df)
        )


        c2.metric(
            "In Progress",
            len(df[df.status=="in progress"])
        )


        c3.metric(
            "1-2 Days Left",
            len(
                df[
                    df.status.isin(
                        [
                            "1 day left",
                            "2 days left"
                        ]
                    )
                ]
            )
        )


        c4.metric(
            "Completed",
            len(df[df.status=="completed"])
        )


        c5.metric(
            "Overdue",
            len(df[df.status=="overdue"])
        )



        # -------------------- TABLE --------------------

        st.subheader("📋 Loan Details")

        st.dataframe(
            filtered_df,
            use_container_width=True
        )



        # -------------------- FINANCIAL --------------------

        st.subheader("💰 Financial Summary")


        f1,f2,f3 = st.columns(3)


        f1.metric(
            "Disbursed",
            f"{df.disbursed_amount.sum():,.0f}"
        )


        f2.metric(
            "Interest",
            f"{df.interest_amount.sum():,.0f}"
        )


        f3.metric(
            "Collection",
            f"{df.collection_amount.sum():,.0f}"
        )



        # -------------------- CHART --------------------

        st.subheader("📊 Status Distribution")

        st.bar_chart(
            df["status"].value_counts()
        )



        # -------------------- DOWNLOAD --------------------

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
            "⬇️ Download Updated Excel",
            to_excel(df),
            file_name="EGSA_short_term_loans_updated.xlsx",
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            )
        )


else:

    st.info(
        "Please upload a file or select Load From GitHub"
    )
