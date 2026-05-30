import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials

SHEET_NAME = "Aaron Attendance System"

def connect_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )

    client = gspread.authorize(creds)
    return client.open(SHEET_NAME)

def get_worksheet(sheet, name, headers):
    try:
        ws = sheet.worksheet(name)
    except:
        ws = sheet.add_worksheet(title=name, rows=1000, cols=len(headers))
        ws.append_row(headers)
    return ws

def attendance_page():
    st.title("Aaron's Attendance Register")

    sheet = connect_sheet()

    workers_ws = get_worksheet(sheet, "Workers", ["Name", "Rate", "BSB", "Account"])
    attendance_ws = get_worksheet(sheet, "Attendance", ["Timestamp", "Date", "Name", "Project", "Action"])

    workers_data = workers_ws.get_all_records()

    if not workers_data:
        st.warning("Please add workers first in the Workers sheet.")
        return

    worker_names = [w["Name"] for w in workers_data]

    name = st.selectbox("Worker Name", worker_names)
    project = st.text_input("Project / Site", value="Carlton")
    action = st.radio("Attendance Type", ["Check In", "Check Out"])

    if st.button("Register Attendance"):
        now = datetime.now()

        attendance_ws.append_row([
            now.strftime("%Y-%m-%d %H:%M:%S"),
            now.strftime("%Y-%m-%d"),
            name,
            project,
            action
        ])

        st.success(f"{action} registered for {name} at {now.strftime('%I:%M %p')}")

def payroll_page():
    st.title("Weekly Payroll Report")

    sheet = connect_sheet()

    workers_ws = get_worksheet(sheet, "Workers", ["Name", "Rate", "BSB", "Account"])
    attendance_ws = get_worksheet(sheet, "Attendance", ["Timestamp", "Date", "Name", "Project", "Action"])

    workers = pd.DataFrame(workers_ws.get_all_records())
    attendance = pd.DataFrame(attendance_ws.get_all_records())

    if attendance.empty:
        st.warning("No attendance records yet.")
        return

    attendance["Timestamp"] = pd.to_datetime(attendance["Timestamp"])
    attendance["Date"] = pd.to_datetime(attendance["Date"])

    today = datetime.now().date()
    start_week = today - timedelta(days=today.weekday())
    end_week = start_week + timedelta(days=6)

    start_date = st.date_input("Start Date", start_week)
    end_date = st.date_input("End Date", end_week)

    filtered = attendance[
        (attendance["Date"].dt.date >= start_date) &
        (attendance["Date"].dt.date <= end_date)
    ]

    results = []

    for name in filtered["Name"].unique():
        worker_records = filtered[filtered["Name"] == name].sort_values("Timestamp")
        total_hours = 0

        for date in worker_records["Date"].dt.date.unique():
            day_records = worker_records[worker_records["Date"].dt.date == date]

            check_in = day_records[day_records["Action"] == "Check In"]["Timestamp"]
            check_out = day_records[day_records["Action"] == "Check Out"]["Timestamp"]

            if not check_in.empty and not check_out.empty:
                start = check_in.iloc[0]
                end = check_out.iloc[-1]

                hours = (end - start).total_seconds() / 3600
                hours -= 0.5
                total_hours += max(hours, 0)

        worker_row = workers[workers["Name"] == name]

        if worker_row.empty:
            rate = 35
        else:
            rate = float(worker_row["Rate"].iloc[0])

        total_pay = total_hours * rate

        results.append({
            "Worker": name,
            "Hours": round(total_hours, 2),
            "Rate": rate,
            "Total Pay": round(total_pay, 2)
        })

    report = pd.DataFrame(results)

    st.dataframe(report, use_container_width=True)

    st.subheader("Payment Message for Aaron")

    message = "Hi Aaron, I’m sending you this week’s payroll details:\n\n"

    for _, row in report.iterrows():
        message += f"{row['Worker']}\n"
        message += f"Hours: {row['Hours']}\n"
        message += f"Rate: ${row['Rate']} per hour\n"
        message += f"Total: ${row['Total Pay']}\n\n"

    st.text_area("Copy this message", message, height=300)
