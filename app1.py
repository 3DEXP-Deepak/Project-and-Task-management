import streamlit as st
import pandas as pd
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode
import io
import openpyxl

# **Streamlit Page Config**
st.set_page_config(page_title="Multi-Project Task Management Dashboard", layout="wide")

# **Define Two Columns for Side-by-Side Layout**
col1, col2 = st.columns([0.5, 1])  # 2:1 ratio for spacing

# **📂 Section 1: File Upload**
with col1:
    st.subheader("📂 Upload Project File")
    uploaded_file = st.file_uploader("Upload your Task Excel File", type=["xlsx"])

# **📊 Section 2: Project Summary**E
with col2:
    st.subheader("📊 Project Summary")

if uploaded_file:
    # **Read Excel File into Memory**
    excel_bytes = uploaded_file.getvalue()
    xl = pd.ExcelFile(io.BytesIO(excel_bytes))
    project_sheets = xl.sheet_names

    # **Sidebar: Create New Project (Placed at Top)**
    st.sidebar.header("📂 Create New Project")
    new_project_name = st.sidebar.text_input("Enter New Project Name")

    if st.sidebar.button("➕ Create New Project"):
        if new_project_name and new_project_name not in project_sheets:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for sheet in project_sheets:
                    xl.parse(sheet).to_excel(writer, sheet_name=sheet, index=False)

                # **Copy tasks from the first available sheet to the new sheet**
                xl.parse(project_sheets[0]).to_excel(writer, sheet_name=new_project_name, index=False)

                writer.close()  # ✅ Fixed Save Error

            output.seek(0)
            xl = pd.ExcelFile(io.BytesIO(output.getvalue()))
            project_sheets.append(new_project_name)
            st.sidebar.success(f"✅ New Project '{new_project_name}' Created!")
        else:
            st.sidebar.error("⚠️ Please enter a unique project name!")

    # **Sidebar: Project Selection & Filters**
    st.sidebar.header("🔍 Select & Filter Tasks")
    selected_project = st.sidebar.selectbox("Select a Project", project_sheets)
    df = xl.parse(selected_project)

    # **Convert Date Columns**
    for col in ["Planned Completion", "Actual Completion", "Planned Start", "Planned End", "Actual End"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    # **Sidebar Filters**
    assignees = df["Assignee"].dropna().unique().tolist()
    statuses = df["Status"].dropna().unique().tolist()

    selected_assignee = st.sidebar.selectbox("Filter by Assignee", ["All"] + assignees)
    selected_status = st.sidebar.selectbox("Filter by Status", ["All"] + statuses)

    # **Apply Filters**
    filtered_df = df.copy()
    if selected_assignee != "All":
        filtered_df = filtered_df[filtered_df["Assignee"] == selected_assignee]
    if selected_status != "All":
        filtered_df = filtered_df[filtered_df["Status"] == selected_status]

    # **📊 Display Project Summary (Placed in col2)**
    with col2:
        st.subheader(f"Project: {selected_project}")
        c1, c2, c3, c4 = st.columns(4)  # Create 4 columns inside col2
        c1.metric("📌 Total Tasks", len(filtered_df))
        c2.metric("✅ Completed Tasks", len(filtered_df[filtered_df["Status"] == "Completed"]))
        c3.metric("⏳ Pending Tasks", len(filtered_df[filtered_df["Status"] == "Pending"]))
        c4.metric("🚀 In Process", len(filtered_df[filtered_df["Status"] == "In process"]))

    # **📅 Planned vs Actual Timeline Chart**
    st.subheader("📅 Planned vs Actual Timeline")
    if "Planned End" in df.columns and "Actual End" in df.columns:
        df["Planned End"] = pd.to_datetime(df["Planned End"])
        df["Actual End"] = pd.to_datetime(df["Actual End"])
        df_melted = df.melt(id_vars=["Task Name"], value_vars=["Planned End", "Actual End"],
                            var_name="Type", value_name="Date")
        color_map_timeline = {"Planned End": "#1f77b4", "Actual End": "#d62728"}
        fig_timeline = px.line(df_melted, x="Date", y="Task Name", color="Type",
                               markers=True, color_discrete_map=color_map_timeline,
                               title="Planned vs Actual Completion")
        st.plotly_chart(fig_timeline, use_container_width=True)

    # **📊 Bar Chart: Task Count by Assignee**
    assignee_task_count = filtered_df["Assignee"].value_counts().reset_index()
    assignee_task_count.columns = ["Assignee", "Task Count"]
    fig_bar = px.bar(assignee_task_count, x="Assignee", y="Task Count", title="Tasks Assigned Per User",
                     color="Task Count", color_continuous_scale="Blues")

    # **📊 Pie Chart: Task Completion Status**
    status_counts = filtered_df["Status"].value_counts().reset_index()
    status_counts.columns = ["Status", "Count"]
    fig_pie = px.pie(status_counts, names="Status", values="Count", title="Task Completion Status",
                      color="Status", color_discrete_map={"Completed": "#2ca02c", "Pending": "#d62728", "In process": "#f1c40f"})

    # **Aligning Charts Side-by-Side**
    col1, col2 = st.columns([2, 1])
    col1.plotly_chart(fig_bar, use_container_width=True)
    col2.plotly_chart(fig_pie, use_container_width=True)

    # **📜 Project Task Data Section**
    st.subheader("📜 Project Task Data")

    gb = GridOptionsBuilder.from_dataframe(filtered_df)
    gb.configure_pagination(enabled=True)
    gb.configure_side_bar()
    gb.configure_selection("multiple", use_checkbox=True)

    # **Editable Columns with Dropdown for "Status" and Date Picker**
    for col in ["Status", "Assignee", "Planned Start", "Planned End", "Actual End", "Delayed", "Comments"]:
        if col in filtered_df.columns:
            if col == "Status":
                gb.configure_column(col, editable=True, cellEditor="agSelectCellEditor",
                                    cellEditorParams={"values": ["Completed", "In process", "Pending"]})
            elif col in ["Planned Start", "Planned End", "Actual End"]:
                gb.configure_column(col, editable=True, cellEditor="agDateCellEditor",
                                    cellEditorParams={"format": "yyyy-MM-dd"})
            else:
                gb.configure_column(col, editable=True)

    grid_options = gb.build()

    grid_response = AgGrid(
        filtered_df, 
        gridOptions=grid_options, 
        fit_columns_on_grid_load=True,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED, 
        update_mode="MANUAL",
        theme="balham",
        height=600
    )

    updated_df = grid_response["data"]

    # **💾 Save Changes to Excel**
    if st.button("💾 Save Changes & Download"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet in project_sheets:
                if sheet == selected_project:
                    updated_df.to_excel(writer, sheet_name=sheet, index=False)
                else:
                    xl.parse(sheet).to_excel(writer, sheet_name=sheet, index=False)

            # **Ensure the New Project Sheet is Added**
            if new_project_name and new_project_name not in xl.sheet_names:
                xl.parse(project_sheets[0]).to_excel(writer, sheet_name=new_project_name, index=False)

            writer.close()

        output.seek(0)
        st.download_button(label="📥 Download Updated Excel", data=output, file_name="updated_tasks.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.success("✅ Changes applied! Download the updated Excel file.")
else:
    st.warning("⚠️ Please upload an Excel file to proceed.")
