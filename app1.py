import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# Set page configuration (wide mode by default)
st.set_page_config(layout="wide")

# Function to load all sheets from an uploaded Excel file
@st.cache_data
def load_all_sheets(file):
    xls = pd.ExcelFile(file)
    sheets_dict = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}
    return sheets_dict

# Streamlit App Title
st.title("📊 VEEMAP Project & Task Management Dashboard")
st.write("Upload an Excel file to analyze task insights across multiple projects.")

# File uploader (Fixed label)
uploaded_file = st.file_uploader("Upload your Task Excel File", type=["xlsx"])

if uploaded_file is not None:
    # Load all sheets into a dictionary of DataFrames
    sheets_dict = load_all_sheets(uploaded_file)

    # Sidebar Filters
    st.sidebar.header("🔍 Filter Projects and Tasks")

    # Dropdown to select a project (sheet name)
    project_names = list(sheets_dict.keys())
    selected_project = st.sidebar.selectbox("Select a Project", project_names)

    # Get the DataFrame for the selected project
    df = sheets_dict[selected_project]

    # Add a "Comments" column if it doesn't exist
    if "Comments" not in df.columns:
        df["Comments"] = ""  # Initialize with empty strings

    # Task Filters
    assignee_filter = st.sidebar.multiselect("Filter by Assignee", df["Assignee"].unique())
    status_filter = st.sidebar.multiselect("Filter by Status", df["Status"].unique())

    # Apply Filters
    filtered_df = df.copy()
    if assignee_filter:
        filtered_df = filtered_df[filtered_df["Assignee"].isin(assignee_filter)]
    if status_filter:
        filtered_df = filtered_df[filtered_df["Status"].isin(status_filter)]

    # Key Metrics
    total_tasks = len(filtered_df)
    completed_tasks = filtered_df[filtered_df["Status"] == "Completed"].shape[0]
    pending_tasks = filtered_df[filtered_df["Status"] != "Completed"].shape[0]
    assignees = filtered_df["Assignee"].nunique()

    # Display Metrics
    st.subheader(f"📊 Task Summary for Project: {selected_project}")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Tasks", total_tasks)
    col2.metric("Completed Tasks", completed_tasks)
    col3.metric("Pending Tasks", pending_tasks)
    col4.metric("Number of Assignees", assignees)

    # Task Completion Rate
    if total_tasks > 0:
        completion_rate = (completed_tasks / total_tasks) * 100
    else:
        completion_rate = 0
    st.metric("Task Completion Rate", f"{completion_rate:.2f}%")

    # Task Completion Status - Bar Chart with Custom Colors
    st.subheader("📈 Task Completion Status")
    if not filtered_df["Status"].empty:
        status_counts = filtered_df["Status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Count"]

        # Updated color mapping
        color_map = {
            "Completed": "#2ca02c",   # Green
            "In process": "#f1c40f",  # Yellow
            "Pending": "#d62728"      # Red
        }

        fig_status = px.bar(status_counts, 
                            x="Status", 
                            y="Count", 
                            color="Status",
                            color_discrete_map=color_map,
                            labels={'x': 'Task Status', 'y': 'Count'}, 
                            title="Task Completion Status")
        st.plotly_chart(fig_status)
    else:
        st.warning("No task data available for visualization.")

    # Task Lists by Status
    st.subheader("📌 Task Status Details")

    # Categorize tasks with assignee names and comments
    completed_tasks_list = filtered_df[filtered_df["Status"] == "Completed"].apply(
        lambda row: f"{row['Task Name']} (Assigned to: {row['Assignee']}) - Comments: {row['Comments']}", axis=1).tolist()

    in_process_tasks_list = filtered_df[filtered_df["Status"] == "In process"].apply(
        lambda row: f"{row['Task Name']} (Assigned to: {row['Assignee']}) - Comments: {row['Comments']}", axis=1).tolist()

    pending_tasks_list = filtered_df[filtered_df["Status"] == "Pending"].apply(
        lambda row: f"{row['Task Name']} (Assigned to: {row['Assignee']}) - Comments: {row['Comments']}", axis=1).tolist()

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("### ✅ Completed Tasks")
        if completed_tasks_list:
            for task in completed_tasks_list:
                st.write(f"- {task}")
        else:
            st.info("No completed tasks.")

    with col2:
        st.markdown("### ⏳ In Process Tasks")
        if in_process_tasks_list:
            for task in in_process_tasks_list:
                st.write(f"- {task}")
        else:
            st.info("No tasks in process.")

    with col3:
        st.markdown("### ❌ Pending Tasks")
        if pending_tasks_list:
            for task in pending_tasks_list:
                st.write(f"- {task}")
        else:
            st.info("No pending tasks.")

    # Add a section for editing comments
    st.subheader("📝 Add/Edit Task Comments")

    # Dropdown to select a task
    task_to_comment = st.selectbox("Select a Task to Add/Edit Comments", filtered_df["Task Name"].unique())

    # Find the selected task's current comment
    current_comment = filtered_df.loc[filtered_df["Task Name"] == task_to_comment, "Comments"].values[0]

    # Text area for comments
    new_comment = st.text_area("Enter your comment", value=current_comment)

    # Save the comment
    if st.button("Save Comment"):
        # Update the comment in the original DataFrame
        sheets_dict[selected_project].loc[sheets_dict[selected_project]["Task Name"] == task_to_comment, "Comments"] = new_comment
        st.success("Comment saved successfully!")

        # Reapply filters to refresh the displayed data
        filtered_df = sheets_dict[selected_project].copy()
        if assignee_filter:
            filtered_df = filtered_df[filtered_df["Assignee"].isin(assignee_filter)]
        if status_filter:
            filtered_df = filtered_df[filtered_df["Status"].isin(status_filter)]

    # Planned vs Actual Timeline - Line Chart
    st.subheader("📅 Planned vs Actual Timeline")
    if "Planned End" in df.columns and "Actual End" in df.columns:
        df["Planned End"] = pd.to_datetime(df["Planned End"])
        df["Actual End"] = pd.to_datetime(df["Actual End"])

        # Melt Data for Line Chart
        df_melted = df.melt(id_vars=["Task Name"], value_vars=["Planned End", "Actual End"],
                            var_name="Type", value_name="Date")

        # Custom colors for Planned vs Actual
        color_map_timeline = {
            "Planned End": "#1f77b4",  # Blue
            "Actual End": "#d62728"    # Red
        }

        fig_timeline = px.line(df_melted, x="Date", y="Task Name", color="Type",
                               markers=True, color_discrete_map=color_map_timeline,
                               labels={"Date": "Timeline", "Task Name": "Tasks"},
                               title="Planned vs Actual Completion")
        st.plotly_chart(fig_timeline)
    else:
        st.warning("Columns for Planned and Actual End Dates not found in uploaded file.")

    # Add a download button for the updated dataset
    st.subheader("💾 Download Updated Data")
    st.write("Click below to download the updated dataset with comments.")
    updated_file = sheets_dict[selected_project].to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download CSV",
        data=updated_file,
        file_name=f"updated_task_data_{selected_project}.csv",
        mime="text/csv"
    )

    # Move Task Data Preview to Last Section
    st.subheader("📋 Task Data Preview")
    st.dataframe(filtered_df)

else:
    st.warning("⚠️ Please upload an Excel file to get started.")