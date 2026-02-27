import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re
from mail import generate_pdf_report,send_email_report
from streamlit_plotly_events import plotly_events

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

            assignee_raw = row_dict.get("Assignee")
            effort = pd.to_numeric(row_dict.get("Effort"), errors="coerce")

            if pd.notna(assignee_raw) and str(assignee_raw).strip() != "" and pd.notna(effort):

                # Convert to string and clean
                assignee_raw = str(assignee_raw).strip()

                # Split multiple names (/, comma)
                assignee_list = re.split(r'[\/,]', assignee_raw)

                for name in assignee_list:

                    name = name.strip()
                    if not name:
                        continue

                    # Take first word only
                    first_name = name.split()[0]

                    # Standardize case (vinay -> Vinay)
                    first_name = first_name.lower().capitalize()

                    final_rows.append({
                        "Ticket Number": row_dict.get("Ticket Number"),
                        "Assignee": first_name,
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
        orientation="h",
        color="Assignee",
        text="Effort",  # 👈 ADD THIS
        color_discrete_sequence=px.colors.qualitative.Set2
    )

    fig_bar.update_traces(
        textposition="outside",  # shows value at end of bar
        textfont_size=12
    )

    fig_bar.update_layout(
        yaxis={'categoryorder': 'total ascending'},
        showlegend=False,
        margin=dict(l=80, r=40, t=40, b=40)
    )

    st.plotly_chart(fig_bar, use_container_width=True)

    # ======================================================
    # MONTHLY TOTAL EFFORT PER CLIENT
    # ======================================================

    st.subheader("Monthly Total Effort per Client")

    monthly_client = (
        filtered_df
        .groupby(["Month", "Month_Num", "Client"], as_index=False)
        .agg({"Effort": "sum"})
        .sort_values("Month_Num")
    )

    # Ensure chronological month ordering
    month_order_sorted = (
        monthly_client[["Month", "Month_Num"]]
        .drop_duplicates()
        .sort_values("Month_Num")["Month"]
        .tolist()
    )

    monthly_client["Month"] = pd.Categorical(
        monthly_client["Month"],
        categories=month_order_sorted,
        ordered=True
    )

    fig_month_client = px.bar(
        monthly_client,
        x="Month",
        y="Effort",
        color="Client",
        barmode="group",
        text="Effort",  # 👈 ADD THIS
        color_discrete_sequence=px.colors.qualitative.Bold
    )

    fig_month_client.update_traces(
        textposition="outside",
        textfont_size=11
    )

    fig_month_client.update_layout(
        xaxis_title="Month",
        yaxis_title="Effort",
        margin=dict(l=60, r=40, t=40, b=40)
    )

    st.plotly_chart(fig_month_client, use_container_width=True)

    # ======================================================
    # CLIENT BREAKDOWN FOR SINGLE ASSIGNEE
    # ======================================================

    if len(assignees) == 1 and len(clients) > 1:
        st.subheader(f"Client-wise Effort Distribution for {assignees[0]}")

        assignee_client = (
            filtered_df
            .groupby("Client", as_index=False)
            .agg({"Effort": "sum"})
            .sort_values("Effort", ascending=False)
        )

        fig_assignee_client = px.bar(
            assignee_client,
            x="Effort",
            y="Client",
            orientation="h",
            color="Client",
            color_discrete_sequence=px.colors.qualitative.Set3
        )

        fig_assignee_client.update_layout(
            showlegend=False,
            yaxis={'categoryorder': 'total ascending'}
        )

        st.plotly_chart(fig_assignee_client, use_container_width=True)

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
        markers=True,
        color_discrete_sequence=px.colors.qualitative.Bold
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
        labels=dict(color="Effort"),
        color_continuous_scale="Viridis"
    )

    st.plotly_chart(fig_heatmap, use_container_width=True)

    # ======================================================
    # DATA TABLE
    # ======================================================

    st.subheader("Filtered Data")
    st.dataframe(filtered_df, use_container_width=True)

    st.divider()
    st.subheader("📧 Email Report")

    receiver_email = st.text_input("Enter recipient email address")

    if st.button("Generate & Send Report"):
        if not receiver_email:
            st.error("Please enter a valid email address.")
        else:
            with st.spinner("Generating report..."):
                kpis = {
                    "effort": int(filtered_df["Effort"].sum()),
                    "tickets": filtered_df["Ticket Number"].nunique(),
                    "assignees": filtered_df["Assignee"].nunique()
                }

                pdf_buffer = generate_pdf_report(
                    fig_bar,
                    fig_month_client,
                    fig_line,
                    fig_heatmap,
                    kpis,
                    clients,
                    assignees,
                    months,
                    selected_weeks,
                    filtered_df  # 👈 pass dataframe
                )
                send_email_report(receiver_email, pdf_buffer)

            st.success("Report sent successfully!")

def normalize_first_name(raw_assignee):

    if pd.isna(raw_assignee):
        return []

    raw_assignee = str(raw_assignee).strip()

    # Split multiple names (/, comma)
    names = re.split(r'[\/,]', raw_assignee)

    cleaned = []

    for name in names:
        name = name.strip()
        if not name:
            continue

        # Take first word only
        first_name = name.split()[0]

        # Standardize case
        first_name = first_name.lower().capitalize()

        cleaned.append(first_name)

    return cleaned