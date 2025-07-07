import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
import streamlit.components.v1 as components

st.set_page_config(page_title="LOS Dashboard", layout="wide")

st.title("Loan Origination System Dashboard")

tabs = st.tabs(["Data","Main Dashboard", "Extras"])

with tabs[0]:
    st.title("Data loading")

    # Load all sheets
    xls = pd.ExcelFile("data cleaned.xlsx")
    df1 = pd.read_excel(xls, sheet_name=0, engine='openpyxl')
    df2 = pd.read_excel(xls, sheet_name=1, engine='openpyxl')
    df3 = pd.read_excel(xls, sheet_name=2, engine='openpyxl')

    st.success("Data loaded from all sheets!")

    df2['CREATEDON'] = pd.to_datetime(df2['CREATEDON'])

    df2['PendingAgingBucket'] = df2.apply(lambda row: (
        "30 - 60 Days" if row['IN_PROGRESS'] == 1 and (datetime.today() - row['CREATEDON']).days <= 60 else
        "60 - 90 Days" if row['IN_PROGRESS'] == 1 and (datetime.today() - row['CREATEDON']).days <= 90 else
        "90 - 120 Days" if row['IN_PROGRESS'] == 1 and (datetime.today() - row['CREATEDON']).days <= 120 else
        "120 - 150 Days" if row['IN_PROGRESS'] == 1 and (datetime.today() - row['CREATEDON']).days <= 150 else
        "Over 150 Days" if row['IN_PROGRESS'] == 1 else
        "N/A"
    ), axis=1)

    st.subheader("Business Hierarchy Preview")
    st.dataframe(df1.head())

    st.subheader("Customer Transaction Log Preview")
    st.dataframe(df2.head())

    st.subheader("Users Preview")
    st.dataframe(df3.head())  

with tabs[1]:
    # KPIs
    total_cases = len(df2)
    in_progress_cases = df2[df2['IN_PROGRESS'] == 1].shape[0]
    approved_cases = df2[df2['WORKFLOWSTEPTYPE_ID'] == 2].shape[0]
    rejected_cases = df2[df2['WORKFLOWSTEPTYPE_ID'] == 3].shape[0]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Cases", total_cases)
    col2.metric("InProgress Cases", in_progress_cases)
    col3.metric("Approved Cases", approved_cases)
    col4.metric("Rejected Cases", rejected_cases)

    # Sort Options
    col_sort1, col_sort2 = st.columns(2)
    with col_sort1:
        dept_sort_order = st.selectbox("Sort Departments by Total Count", ["Descending", "Ascending"], key="dept_sort")
    with col_sort2:
        user_sort_order = st.selectbox("Sort Users by Total Count", ["Descending", "Ascending"], key="user_sort")

    status_map = {2: "Approved", 3: "Rejected"}

    # --- Department Chart ---
    df_status_dept = pd.merge(
        df2[['TRANSACTION_ID', 'BUSINESSHIERARCHY_ID', 'WORKFLOWSTEPTYPE_ID']],
        df1[['BUSINESSHIERARCHY_ID', 'TITLE']],
        on='BUSINESSHIERARCHY_ID', how='left'
    )
    df_status_dept['Case Status'] = df_status_dept['WORKFLOWSTEPTYPE_ID'].map(status_map).fillna("In Progress")
    df_dept_counts = df_status_dept.groupby(['TITLE', 'Case Status'])['TRANSACTION_ID'].count().reset_index()

    dept_totals = df_dept_counts.groupby('TITLE')['TRANSACTION_ID'].sum().reset_index(name='Total')
    dept_sorted = dept_totals.sort_values('Total', ascending=(dept_sort_order == "Ascending"))['TITLE'].tolist()
    df_dept_counts['TITLE'] = pd.Categorical(df_dept_counts['TITLE'], categories=dept_sorted, ordered=True)

    fig_dept = px.bar(
        df_dept_counts, x="TRANSACTION_ID", y="TITLE", color="Case Status", orientation="h",
        title="Count of TRANSACTION_ID by TITLE and Case Status",
        category_orders={"TITLE": dept_sorted},
        color_discrete_map={"Approved": "#1f77b4", "In Progress": "#ff7f0e", "Rejected": "#d62728"}
    )
    fig_dept.update_layout(height=20 * len(dept_sorted), margin=dict(l=100, r=30, t=50, b=30),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font=dict(color='white'))
    fig_dept.update_yaxes(automargin=True, tickfont=dict(size=12), color='white')
    fig_dept.update_xaxes(color='white')

    # --- User Chart ---
    df_status_user = pd.merge(
        df2[['TRANSACTION_ID', 'CREATEDBY', 'WORKFLOWSTEPTYPE_ID']],
        df3[['USER_ID', 'NAME']],
        left_on='CREATEDBY', right_on='USER_ID', how='left'
    )
    df_status_user['Case Status'] = df_status_user['WORKFLOWSTEPTYPE_ID'].map(status_map).fillna("In Progress")
    df_user_counts = df_status_user.groupby(['NAME', 'Case Status'])['TRANSACTION_ID'].count().reset_index()

    user_totals = df_user_counts.groupby("NAME")["TRANSACTION_ID"].sum().reset_index(name="Total")
    user_sorted = user_totals.sort_values("Total", ascending=(user_sort_order == "Ascending"))["NAME"].tolist()
    df_user_counts["NAME"] = pd.Categorical(df_user_counts["NAME"], categories=user_sorted, ordered=True)

    fig_user = px.bar(
        df_user_counts, x="TRANSACTION_ID", y="NAME", color="Case Status", orientation="h",
        title="Count of TRANSACTION_ID by NAME and Case Status",
        category_orders={"NAME": user_sorted},
        color_discrete_map={"Approved": "#1f77b4", "In Progress": "#ff7f0e", "Rejected": "#d62728"}
    )
    fig_user.update_layout(height=20 * len(user_sorted), margin=dict(l=0, r=30, t=30, b=20),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font=dict(color='white'))
    fig_user.update_yaxes(automargin=True, tickfont=dict(size=12), color='white')
    fig_user.update_xaxes(color='white')
    user_html = fig_user.to_html(full_html=False, include_plotlyjs='cdn')

    # Build Aging Matrices
    df2_trimmed = df2[['BUSINESSHIERARCHY_ID', 'CREATEDBY', 'TRANSACTION_ID', 'IN_PROGRESS', 'CREATEDON', 'PendingAgingBucket']]
    df_merged = pd.merge(df2_trimmed[df2_trimmed['IN_PROGRESS'] == 1], df1[['BUSINESSHIERARCHY_ID', 'TITLE']], on='BUSINESSHIERARCHY_ID', how='left')

    matrix = df_merged.groupby(['TITLE', 'PendingAgingBucket'])['TRANSACTION_ID'].count().reset_index()
    matrix = matrix.pivot(index='TITLE', columns='PendingAgingBucket', values='TRANSACTION_ID').fillna(0).astype(int)
    desired_order = ["30 - 60 Days", "60 - 90 Days", "90 - 120 Days", "120 - 150 Days", "Over 150 Days"]
    available_columns = [col for col in desired_order if col in matrix.columns]
    Business_aging_matrix = matrix[available_columns]
    Business_aging_matrix['Total'] = Business_aging_matrix.sum(axis=1)

    df_merged2 = pd.merge(df2_trimmed[df2_trimmed['IN_PROGRESS'] == 1], df3[['USER_ID','NAME']], left_on='CREATEDBY', right_on='USER_ID', how='left')
    matrix2 = df_merged2.groupby(['NAME', 'PendingAgingBucket'])['TRANSACTION_ID'].count().reset_index()
    matrix2 = matrix2.pivot(index='NAME', columns='PendingAgingBucket', values='TRANSACTION_ID').fillna(0).astype(int)
    available_columns = [col for col in desired_order if col in matrix2.columns]
    User_aging_matrix = matrix2[available_columns]
    User_aging_matrix['Total'] = User_aging_matrix.sum(axis=1)

    # Display Charts and Matrices
    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
        st.subheader("Case Status by Department")
        st.plotly_chart(fig_dept, use_container_width=True)
        st.subheader("Aging Matrix by Department")
        st.dataframe(Business_aging_matrix.reset_index(), use_container_width=True)

    with col_chart2:
        st.subheader("Case Status by User")
        components.html(
            f"""
            <div style="height:400px; overflow-y:auto; padding:10px">
            {user_html}
            </div>
            """,
            height=450,
            scrolling=True
        )   

        st.subheader("Aging Matrix by User")
        st.dataframe(User_aging_matrix.reset_index(), use_container_width=True)

with tabs[2]:
    st.title("Extras Dashboard")

    # --- KPI Calculations ---
    revert_cases = df2[df2['WORKFLOWSTEPTYPE_ID'] == 4].shape[0]
    forward_cases = df2[df2['WORKFLOWSTEPTYPE_ID'] == 5].shape[0]
    null_cases = df2[df2['WORKFLOWSTEPTYPE_ID'].isnull() | (df2['WORKFLOWSTEPTYPE_ID'] == 0)].shape[0]

    col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
    col_kpi1.metric("Revert Cases", revert_cases)
    col_kpi2.metric("Null Cases", null_cases)
    col_kpi3.metric("Forward Cases", forward_cases)

    # --- Year Filter ---
    df2['CREATEDON'] = pd.to_datetime(df2['CREATEDON'], errors='coerce')
    df2['Year'] = df2['CREATEDON'].dt.year
    selected_years = st.multiselect("Select Year(s)", sorted(df2['Year'].dropna().unique()), default=sorted(df2['Year'].dropna().unique()))

    filtered_df = df2[df2['Year'].isin(selected_years)]

    # --- Line Chart: Transaction Volume Over Time ---
    volume_over_time = filtered_df.groupby(pd.Grouper(key='CREATEDON', freq='M'))['TRANSACTION_ID'].count().reset_index()
    volume_over_time.columns = ['Time', 'Number of Cases']

    fig_line = px.line(volume_over_time, x='Time', y='Number of Cases', title='Transaction Volume Over Time')
    fig_line.update_traces(mode="lines+markers")
    fig_line.update_layout(
        height=400,
        margin=dict(l=40, r=20, t=50, b=40),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white')
    )
    fig_line.update_xaxes(color='white')
    fig_line.update_yaxes(color='white')

    # --- Donut Chart: Volume by Department ---
    df_dept_volume = pd.merge(
        filtered_df[['TRANSACTION_ID', 'BUSINESSHIERARCHY_ID']],
        df1[['BUSINESSHIERARCHY_ID', 'TITLE']],
        on='BUSINESSHIERARCHY_ID', how='left'
    )
    dept_counts = df_dept_volume.groupby('TITLE')['TRANSACTION_ID'].count().reset_index()
    dept_counts = dept_counts.sort_values('TRANSACTION_ID', ascending=False)

    fig_donut = px.pie(
        dept_counts,
        names='TITLE',
        values='TRANSACTION_ID',
        hole=0.5,
        title='Transaction Volume by Department',
    )
    fig_donut.update_layout(
        height=400,
        margin=dict(l=40, r=20, t=50, b=40),
        showlegend=True,
        font=dict(color='white'),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
    )

    # --- Display Side-by-Side ---
    st.plotly_chart(fig_line, use_container_width=True)
    st.plotly_chart(fig_donut, use_container_width=True)

