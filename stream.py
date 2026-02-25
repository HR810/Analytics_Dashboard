import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re

st.set_page_config(layout="wide")

# ======================================================
# DATA LOADING (CACHED)
# ======================================================

@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
    final_rows = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        current_month = None
        current_week = None
        columns = None

        for _, row in df.iterrows():
            row_str = row.fillna("").astype(str).str.strip()
            row_text = " ".join(row_str)

            # Detect Month
            month_match = re.search(
                r'\b(January|February|March|April|May|June|July|August|September|October|November|December)\b',
                row_text,
                re.IGNORECASE
            )
            if month_match:
                current_month = month_match.group().capitalize()
                continue

            # Detect Week
            week_match = re.search(r'\bWeek\s*\d+\b', row_text, re.IGNORECASE)
            if week_match:
                current_week = re.sub(r'\s+', '', week_match.group()).capitalize()
                continue

            # Detect Header
            if "Ticket Number" in row_str.values:
                columns = row_str.tolist()
                continue

            if columns is None:
                continue

            row_dict = dict(zip(columns, row.tolist()))

            assignee = row_dict.get("Assignee")
            effort = pd.to_numeric(row_dict.get("Effort"), errors="coerce")

            if pd.notna(assignee) and str(assignee).strip() != "" and pd.notna(effort):
                final_rows.append({
                    "Ticket Number": row_dict.get("Ticket Number"),
                    "Assignee": str(assignee).strip(),
                    "Effort": effort,
                    "Month": current_month,
                    "Week": current_week,
                    "Client": sheet_name
                })

    final_df = pd.DataFrame(final_rows)

    month_order = {
        "January": 1, "February": 2, "March": 3,
        "April": 4, "May": 5, "June": 6,
        "July": 7, "August": 8, "September": 9,
        "October": 10, "November": 11, "December": 12
    }

    final_df["Month_Num"] = final_df["Month"].map(month_order)
    final_df["Week_Num"] = final_df["Week"].str.extract(r'(\d+)').astype(int)
    final_df["Time_Order"] = final_df["Month_Num"] * 10 + final_df["Week_Num"]
    final_df["Time_Label"] = final_df["Month"] + " " + final_df["Week"]

    return final_df


# ======================================================
# UI
# ======================================================

st.title("📊 Customer Ticket Effort Dashboard")

uploaded_file = st.file_uploader("Upload Customer_Ticket_Status.xlsx", type=["xlsx"])

if uploaded_file:
    df = load_data(uploaded_file)

    # ======================================================
    # SIDEBAR FILTERS
    # ======================================================

    st.sidebar.header("Filters")

    clients = st.sidebar.multiselect(
        "Select Client",
        options=sorted(df["Client"].unique()),
        default=sorted(df["Client"].unique())
    )

    assignees = st.sidebar.multiselect(
        "Select Assignee",
        options=sorted(df["Assignee"].unique()),
        default=sorted(df["Assignee"].unique())
    )

    months = st.sidebar.multiselect(
        "Select Month",
        options=sorted(
            df["Month"].dropna().unique(),
            key=lambda x: df[df["Month"] == x]["Month_Num"].iloc[0]
        ),
        default=sorted(
            df["Month"].dropna().unique(),
            key=lambda x: df[df["Month"] == x]["Month_Num"].iloc[0]
        )
    )

    # Build chronological week list AFTER month selection
    week_timeline = (
        df[df["Month"].isin(months)]
        [["Time_Order", "Time_Label"]]
        .drop_duplicates()
        .sort_values("Time_Order")
    )

    week_order = week_timeline["Time_Label"].tolist()

    selected_weeks = st.sidebar.multiselect(
        "Select Week",
        options=week_order,
        default=week_order
    )

    top_n = st.sidebar.slider("Show Top N Clients (by total effort)", 3, 20, 8)

    # ======================================================
    # FILTER DATA
    # ======================================================

    filtered_df = df[
        (df["Client"].isin(clients)) &
        (df["Assignee"].isin(assignees)) &
        (df["Month"].isin(months)) &
        (df["Time_Label"].isin(selected_weeks))
    ]

    if filtered_df.empty:
        st.warning("No data available for selected filters.")
        st.stop()

    # ======================================================
    # KPI SECTION
    # ======================================================

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Effort", int(filtered_df["Effort"].sum()))
    col2.metric("Total Tickets", filtered_df["Ticket Number"].nunique())
    col3.metric("Active Assignees", filtered_df["Assignee"].nunique())

    st.divider()

    # ======================================================
    # MONTHLY ASSIGNEE BAR CHART
    # ======================================================

    st.subheader("Monthly Total Effort per Assignee")

    monthly_assignee = (
        filtered_df
        .groupby("Assignee", as_index=False)
        .agg({"Effort": "sum"})
        .sort_values("Effort", ascending=False)
    )

    fig_bar = px.bar(
        monthly_assignee,
        x="Effort",
        y="Assignee",
        orientation="h"
    )

    fig_bar.update_layout(yaxis={'categoryorder': 'total ascending'})

    st.plotly_chart(fig_bar, use_container_width=True)

    # ======================================================
    # WEEKLY TREND PER CLIENT
    # ======================================================

    st.subheader("Weekly Effort Trend per Client")

    weekly = (
        filtered_df
        .groupby(["Time_Order", "Time_Label", "Client"], as_index=False)
        .agg({"Effort": "sum"})
        .sort_values("Time_Order")
    )

    # Top N clients
    top_clients = (
        filtered_df.groupby("Client")["Effort"]
        .sum()
        .sort_values(ascending=False)
        .head(top_n)
        .index
    )

    weekly = weekly[weekly["Client"].isin(top_clients)]

    # Force chronological categorical ordering
    weekly["Time_Label"] = pd.Categorical(
        weekly["Time_Label"],
        categories=week_order,
        ordered=True
    )

    fig_line = px.line(
        weekly.sort_values("Time_Order"),
        x="Time_Label",
        y="Effort",
        color="Client",
        markers=True
    )

    fig_line.update_layout(
        xaxis=dict(
            title="Week",
            categoryorder="array",
            categoryarray=week_order
        ),
        yaxis_title="Effort",
        hovermode="x unified"
    )

    st.plotly_chart(fig_line, use_container_width=True)

    # ======================================================
    # WEEKLY HEATMAP
    # ======================================================

    st.subheader("Weekly Effort Heatmap per Client")

    pivot = (
        weekly
        .pivot_table(
            values="Effort",
            index="Client",
            columns="Time_Label",
            fill_value=0
        )
        .reindex(columns=week_order)
    )

    fig_heatmap = px.imshow(
        pivot,
        aspect="auto",
        labels=dict(color="Effort")
    )

    st.plotly_chart(fig_heatmap, use_container_width=True)

    # ======================================================
    # DATA TABLE
    # ======================================================

    st.subheader("Filtered Data")
    st.dataframe(filtered_df, use_container_width=True)