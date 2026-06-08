import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from math import ceil
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

try:
    from streamlit_js_eval import get_geolocation
except Exception:
    get_geolocation = None

# =========================================================
# CONFIG
# =========================================================

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQi2XaKTk3vZnIWejQCiCVNqLwJRcDKYThJCKHOH4iPA_JdDQFEUTcKq5BYRDtbAFDYcu6gWnqQgH2/pub?output=csv"
SPREADSHEET_NAME = "Aarons Pricing Database"
COSTING_SHEET_NAME = "Costing_Records"

MELBOURNE_TZ = ZoneInfo("Australia/Melbourne")

def melbourne_now():
    """Current Melbourne time, used for all attendance records."""
    return datetime.now(MELBOURNE_TZ).replace(tzinfo=None)

EXPECTED_HEADERS = [
    "quote_id",
    "date",
    "client_name",
    "project_address",
    "project_type",
    "working_hours",
    "floor_level",
    "room_area",
    "item",
    "category",
    "unit",
    "measurement_method",
    "quantity",
    "rate",
    "item_total",
    "worker_days",
    "equipment_cost",
    "disposal_cost",
    "truck_cost",
    "other_cost",
    "difficulty_percent",
    "difficulty_allowance",
    "labour_cost",
    "supervisor_cost",
    "direct_cost",
    "recommended_price",
    "gst_total",
    "final_total",
    "notes",
]

st.set_page_config(
    page_title="Aaron's Pricing Calculator",
    page_icon="🧾",
    layout="wide"
)


# =========================================================
# ATTENDANCE & PAYROLL HELPERS
# =========================================================

WORKERS_SHEET_NAME = "Workers"
ATTENDANCE_SHEET_NAME = "Attendance"
PROJECTS_SHEET_NAME = "Projects"

# Workers register bank details once. Rates can be edited manually later in Google Sheets.
WORKERS_HEADERS = ["Name", "Rate", "BSB", "Account", "Sunday Rate", "Active", "Email"]
ATTENDANCE_HEADERS = ["Timestamp", "Date", "Name", "Project", "Action", "Latitude", "Longitude", "Location Link"]
PROJECTS_HEADERS = ["Project", "Active"]
ATTENDANCE_ACTIONS = ["Check In", "Lunch Start", "Lunch End", "Check Out"]


@st.cache_resource(show_spinner=False)
def connect_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )

    client = gspread.authorize(creds)
    return client.open(SPREADSHEET_NAME)


@st.cache_resource(show_spinner=False)
def get_worksheet_cached(sheet_name, headers_tuple):
    """Cache worksheet objects so the app does not hit Google Sheets metadata on every page load.
    IMPORTANT: Create the tabs once with the correct headers; this function only creates them if missing.
    """
    spreadsheet = connect_spreadsheet()
    headers = list(headers_tuple)
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(
            title=sheet_name,
            rows="3000",
            cols=str(len(headers) + 5)
        )
        ws.append_row(headers, value_input_option="RAW")
        return ws


def get_or_create_worksheet(spreadsheet, sheet_name, headers):
    """Backward-compatible wrapper. Uses cached worksheet objects to avoid Google Sheets 429 quota errors."""
    return get_worksheet_cached(sheet_name, tuple(headers))


@st.cache_data(ttl=300, show_spinner=False)
def get_sheet_records_cached(sheet_name, headers_tuple):
    """Read records from Google Sheets with a short cache to avoid Google API 429 quota errors."""
    spreadsheet = connect_spreadsheet()
    ws = get_or_create_worksheet(spreadsheet, sheet_name, list(headers_tuple))
    return ws.get_all_records()


def clear_sheet_cache():
    """Clear cached sheet reads after writing new rows."""
    try:
        st.cache_data.clear()
    except Exception:
        pass


def get_project_options(projects_ws):
    projects = get_sheet_records_cached(PROJECTS_SHEET_NAME, tuple(PROJECTS_HEADERS))
    active_projects = []

    for project in projects:
        project_name = clean_text(project.get("Project", ""))
        active = clean_text(project.get("Active", "Yes")).lower()
        if project_name and active not in ["no", "inactive", "false", "0"]:
            active_projects.append(project_name)

    return sorted(set(active_projects))


def project_picker(projects_ws):
    project_options = get_project_options(projects_ws)

    if project_options:
        selected = st.selectbox("Project / Site", project_options)
        other_project = st.text_input("Other project, if not listed", value="")
        return other_project.strip() if other_project.strip() else selected

    st.info("No projects found yet. Add projects in the Projects sheet or type the project manually below.")
    return st.text_input("Project / Site", value="Carlton")


def clean_text(value):
    return str(value).strip()


def safe_float(value, default=0.0):
    try:
        if value is None or str(value).strip() == "":
            return default
        return float(str(value).replace("$", "").replace(",", "").strip())
    except Exception:
        return default


def normalize_bsb(value):
    """Keep BSB as 6-digit text, including leading zero."""
    cleaned = "".join(ch for ch in str(value).strip() if ch.isdigit())
    if cleaned == "":
        return ""
    return cleaned.zfill(6)


def normalize_account(value):
    """Keep account number as text."""
    return "".join(ch for ch in str(value).strip() if ch.isdigit())



def get_current_location():
    """Ask the phone/browser for GPS location and return lat, lon and Google Maps link."""
    if get_geolocation is None:
        return "", "", ""

    try:
        location = get_geolocation()
        if not location or "coords" not in location:
            return "", "", ""

        coords = location.get("coords", {})
        lat = coords.get("latitude", "")
        lon = coords.get("longitude", "")

        if lat == "" or lon == "":
            return "", "", ""

        maps_link = f"https://www.google.com/maps?q={lat},{lon}"
        return lat, lon, maps_link
    except Exception:
        return "", "", ""



def location_required_or_stop(lat, lon, location_link):
    """Block attendance registration if location is not detected."""
    if not lat or not lon or not location_link:
        st.error(
            "Location not detected. Please allow location access on your phone/browser, "
            "refresh this page, and try again."
        )
        st.info(
            "How to fix it: tap the lock/info icon next to the website address, "
            "open Site settings or Permissions, set Location to Allow, then refresh."
        )
        if st.button("Refresh location / Try again", use_container_width=True):
            st.rerun()
        st.stop()


def make_payroll_excel(summary_df, daily_df, selected_label):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Payroll Summary", index=False, startrow=3)
        daily_df.to_excel(writer, sheet_name="Daily Attendance Detail", index=False)

        ws = writer.book["Payroll Summary"]
        ws["A1"] = "Aaron's Demolition - Weekly Payroll Report"
        ws["A2"] = f"Week: {selected_label}"

        for sheet_name in writer.book.sheetnames:
            ws2 = writer.book[sheet_name]
            for col in ws2.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        max_len = max(max_len, len(str(cell.value)))
                    except Exception:
                        pass
                ws2.column_dimensions[col_letter].width = min(max_len + 2, 35)

    output.seek(0)
    return output.getvalue()


def clean_report_for_email(report_df):
    df = report_df.copy()
    if "Email" in df.columns:
        df = df.drop(columns=["Email"])
    return df


def dataframe_to_email_table(df):
    if df is None or df.empty:
        return "<p>No records found.</p>"

    html = df.to_html(
        index=False,
        border=0,
        classes="payroll-table",
        escape=False
    )

    # Remove line breaks/extra indentation so Streamlit preview does not show raw HTML/code blocks.
    html = " ".join(str(html).split())
    return html


def email_css():
    return """
<style>
body{font-family:Arial,sans-serif;font-size:14px;color:#222;line-height:1.45;}
h2{color:#1f2937;margin-top:22px;margin-bottom:10px;font-size:18px;}
h3{color:#374151;margin-top:18px;margin-bottom:8px;font-size:15px;}
.summary-box{background:#f3f4f6;border:1px solid #d1d5db;border-radius:8px;padding:12px;margin:12px 0 18px 0;}
table.payroll-table{border-collapse:collapse;width:100%;margin-bottom:18px;}
table.payroll-table th{background:#1f2937;color:#fff;padding:8px;border:1px solid #d1d5db;text-align:left;font-size:12px;}
table.payroll-table td{padding:7px;border:1px solid #d1d5db;font-size:12px;vertical-align:top;}
table.payroll-table tr:nth-child(even){background:#f9fafb;}
</style>
"""


def build_attendance_detail_html(daily_df):
    if daily_df is None or daily_df.empty:
        return "<p>No daily attendance detail available.</p>"

    detail_cols = [
        "Date",
        "Project",
        "Check In",
        "Lunch Start",
        "Lunch End",
        "Check Out",
        "Lunch Hours",
        "Payable Hours",
        "Note"
    ]

    parts = []

    for worker in sorted(daily_df["Worker"].dropna().astype(str).unique()):
        worker_df = daily_df[daily_df["Worker"].astype(str) == worker].copy()
        cols = [c for c in detail_cols if c in worker_df.columns]
        parts.append(f"<h3>{worker}</h3>")
        parts.append(dataframe_to_email_table(worker_df[cols]))

    return "".join(parts)


def build_payroll_email_html(report_df, daily_df, selected_label, total_hours, total_pay, custom_message=""):
    report_df = clean_report_for_email(report_df)
    email_df = report_df.copy()

    payroll_cols = [
        "Name",
        "BSB",
        "Account",
        "Normal Hours",
        "Sunday Hours",
        "Total Hours",
        "Lunch Deducted",
        "Rate",
        "Sunday Rate",
        "Total Pay",
        "Adjustment AUD"
    ]
    payroll_cols = [c for c in payroll_cols if c in email_df.columns]
    email_df = email_df[payroll_cols]

    for col in ["Normal Hours", "Sunday Hours", "Total Hours", "Lunch Deducted", "Rate", "Sunday Rate", "Adjustment AUD", "Total Pay"]:
        if col in email_df.columns:
            email_df[col] = email_df[col].map(lambda x: f"{float(x):,.2f}" if str(x) != "" else "")

    payroll_table = dataframe_to_email_table(email_df)
    attendance_detail = build_attendance_detail_html(daily_df)

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
{email_css()}
</head>
<body>
<p>Hi Aaron,</p>
<p>Please find below the payroll summary for the selected period <b>{selected_label}</b>.</p>
<h2>Payroll Summary</h2>
{payroll_table}
<div class="summary-box">
<b>Summary</b><br>
Total Workers: {len(report_df)}<br>
Total Payable Hours: {total_hours:,.2f}<br>
Total Payroll: AUD {total_pay:,.2f}
</div>
<h2>Daily Attendance Detail</h2>
{attendance_detail}
<p>The detailed payroll report is attached for your records.</p>
<p>Kind regards,<br>Lionel Castro<br>Aaron's Demolition</p>
</body>
</html>"""
    return html.strip()


def build_worker_email_html(worker_name, worker_summary_df, worker_daily_df, selected_label, custom_message=""):
    summary_df = clean_report_for_email(worker_summary_df)

    summary_cols = [
        "Name",
        "Normal Hours",
        "Sunday Hours",
        "Total Hours",
        "Lunch Deducted",
        "Rate",
        "Sunday Rate",
        "Total Pay",
        "Adjustment AUD"
    ]
    summary_cols = [c for c in summary_cols if c in summary_df.columns]
    summary_df = summary_df[summary_cols]

    for col in ["Normal Hours", "Sunday Hours", "Total Hours", "Lunch Deducted", "Rate", "Sunday Rate", "Adjustment AUD", "Total Pay"]:
        if col in summary_df.columns:
            summary_df[col] = summary_df[col].map(lambda x: f"{float(x):,.2f}" if str(x) != "" else "")

    detail_cols = [
        "Date",
        "Project",
        "Check In",
        "Lunch Start",
        "Lunch End",
        "Check Out",
        "Lunch Hours",
        "Payable Hours",
        "Note"
    ]
    detail_cols = [c for c in detail_cols if c in worker_daily_df.columns]

    first_name = str(worker_name).split()[0] if str(worker_name).strip() else "there"

    summary_table = dataframe_to_email_table(summary_df)
    detail_table = dataframe_to_email_table(worker_daily_df[detail_cols]) if not worker_daily_df.empty else "<p>No daily attendance detail available.</p>"

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
{email_css()}
</head>
<body>
<p>Hi {first_name},</p>
<p>Please find below your attendance and payroll summary for the selected period <b>{selected_label}</b>.</p>
<h2>Payroll Summary</h2>
{summary_table}
<h2>Daily Attendance Detail</h2>
{detail_table}
<p>If you have any questions about your attendance records, please contact management.</p>
<p>Kind regards,<br>Lionel Castro<br>Aaron's Demolition</p>
</body>
</html>"""
    return html.strip()

def get_secret_value(key, default=""):
    # First checks normal top-level Streamlit secrets.
    value = st.secrets.get(key, default)

    # If email secrets were pasted after [gcp_service_account], TOML stores them inside that block.
    if (value == default or value is None or value == "") and "gcp_service_account" in st.secrets:
        value = st.secrets["gcp_service_account"].get(key, default)

    return value


def send_payroll_email(to_email, subject, html_body, attachment_bytes=None, attachment_filename=None):
    sender = str(get_secret_value("EMAIL_SENDER", "")).strip()
    password = str(get_secret_value("EMAIL_PASSWORD", "")).replace(" ", "").strip()
    smtp_server = str(get_secret_value("SMTP_SERVER", "smtp.gmail.com")).strip()
    smtp_port = int(get_secret_value("SMTP_PORT", 587))

    if not sender or not password:
        raise ValueError(
            "Missing EMAIL_SENDER or EMAIL_PASSWORD in Streamlit Secrets. "
            "Add them as top-level keys, or the app will also read them from [gcp_service_account] if pasted there."
        )

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(html_body, "html"))

    if attachment_bytes is not None and attachment_filename:
        attachment = MIMEApplication(attachment_bytes, _subtype="xlsx")
        attachment.add_header("Content-Disposition", "attachment", filename=attachment_filename)
        msg.attach(attachment)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)

def attendance_page():
    st.title("👷 Aaron's Attendance Register")
    st.caption("Scan the QR, register once, then always mark attendance under your own name.")

    try:
        spreadsheet = connect_spreadsheet()
        workers_ws = get_or_create_worksheet(spreadsheet, WORKERS_SHEET_NAME, WORKERS_HEADERS)
        attendance_ws = get_or_create_worksheet(spreadsheet, ATTENDANCE_SHEET_NAME, ATTENDANCE_HEADERS)
        projects_ws = get_or_create_worksheet(spreadsheet, PROJECTS_SHEET_NAME, PROJECTS_HEADERS)
    except Exception as e:
        st.error("Could not connect to Google Sheets.")
        st.exception(e)
        return

    st.markdown("### Location")
    lat, lon, location_link = get_current_location()

    if location_link:
        st.success("Location detected automatically.")
        st.caption(location_link)
    else:
        location_required_or_stop(lat, lon, location_link)

    workers_data = get_sheet_records_cached(WORKERS_SHEET_NAME, tuple(WORKERS_HEADERS))
    worker_names = [clean_text(w.get("Name", "")) for w in workers_data if clean_text(w.get("Name", ""))]

    register_type = st.radio(
        "Choose one option",
        ["I am already registered", "First time registration"],
        horizontal=True
    )

    project = project_picker(projects_ws)

    if register_type == "First time registration":
        st.markdown("### Worker Details")
        st.info("Your BSB, bank account and email are saved only once. Next time, choose 'I am already registered'.")

        first_name = st.text_input("First Name")
        last_name = st.text_input("Last Name")
        email = st.text_input("Email")
        bsb = st.text_input("BSB")
        account = st.text_input("Account Number")

        st.info("First time registration will automatically create a Check In record.")

        if st.button("✅ Register Worker and Check In", use_container_width=True):
            location_required_or_stop(lat, lon, location_link)

            full_name = f"{first_name.strip()} {last_name.strip()}".strip()
            bsb_clean = normalize_bsb(bsb)
            account_clean = normalize_account(account)
            email_clean = clean_text(email)

            if not first_name.strip() or not last_name.strip() or not bsb_clean or not account_clean or not email_clean:
                st.warning("Please complete first name, last name, email, BSB and account number.")
                return

            now = melbourne_now()

            if full_name not in worker_names:
                workers_ws.append_row([
                    full_name,
                    35,
                    bsb_clean,
                    account_clean,
                    0,
                    "Yes",
                    email_clean
                ], value_input_option="RAW")
                clear_sheet_cache()
            else:
                st.info("This worker already exists. Bank details were not duplicated.")

            attendance_ws.append_row([
                now.strftime("%Y-%m-%d %H:%M:%S"),
                now.strftime("%Y-%m-%d"),
                full_name,
                project,
                "Check In",
                lat,
                lon,
                location_link
            ], value_input_option="RAW")

            clear_sheet_cache()
            st.success(f"Check In registered for {full_name} at {now.strftime('%I:%M %p')}.")

    else:
        if not worker_names:
            st.warning("No registered workers found yet. Please choose 'First time registration'.")
            return

        name = st.selectbox("Worker Name", sorted(worker_names))
        action = st.radio(
            "Attendance Type",
            ATTENDANCE_ACTIONS,
            horizontal=True,
            key="existing_worker_action"
        )

        if st.button("✅ Register Attendance", use_container_width=True):
            location_required_or_stop(lat, lon, location_link)

            now = melbourne_now()

            attendance_ws.append_row([
                now.strftime("%Y-%m-%d %H:%M:%S"),
                now.strftime("%Y-%m-%d"),
                name,
                project,
                action,
                lat,
                lon,
                location_link
            ], value_input_option="RAW")

            clear_sheet_cache()
            st.success(f"{action} registered for {name} at {now.strftime('%I:%M %p')}.")


def get_week_options(attendance_dates):
    today = melbourne_now().date()
    current_monday = today - timedelta(days=today.weekday())

    valid_dates = [d.date() for d in attendance_dates.dropna()] if not attendance_dates.empty else []
    if valid_dates:
        first_date = min(valid_dates)
        first_monday = first_date - timedelta(days=first_date.weekday())
    else:
        first_monday = current_monday

    weeks = []
    cursor = current_monday
    while cursor >= first_monday:
        sunday = cursor + timedelta(days=6)
        label = f"{cursor.strftime('%d/%m/%Y')} - {sunday.strftime('%d/%m/%Y')}"
        weeks.append((label, cursor, sunday))
        cursor -= timedelta(days=7)

    return weeks


def payroll_page():
    st.title("💵 Weekly Payroll Dashboard")
    st.caption("Weeks run from Monday to Sunday. Select the week, review hours, BSB, account number and total payment.")

    try:
        spreadsheet = connect_spreadsheet()
        workers_ws = get_or_create_worksheet(spreadsheet, WORKERS_SHEET_NAME, WORKERS_HEADERS)
        attendance_ws = get_or_create_worksheet(spreadsheet, ATTENDANCE_SHEET_NAME, ATTENDANCE_HEADERS)
        projects_ws = get_or_create_worksheet(spreadsheet, PROJECTS_SHEET_NAME, PROJECTS_HEADERS)
    except Exception as e:
        st.error("Could not connect to Google Sheets.")
        st.exception(e)
        return

    workers = pd.DataFrame(get_sheet_records_cached(WORKERS_SHEET_NAME, tuple(WORKERS_HEADERS)))
    attendance = pd.DataFrame(get_sheet_records_cached(ATTENDANCE_SHEET_NAME, tuple(ATTENDANCE_HEADERS)))

    if workers.empty:
        st.warning("No workers found. Workers must register first from the QR page.")
        return

    if attendance.empty:
        st.warning("No attendance records yet.")
        return

    required_cols = ["Timestamp", "Date", "Name", "Project", "Action"]
    missing_cols = [c for c in required_cols if c not in attendance.columns]
    if missing_cols:
        st.error(f"Missing columns in Attendance sheet: {missing_cols}")
        return

    attendance["Timestamp"] = pd.to_datetime(attendance["Timestamp"], errors="coerce")
    attendance["Date"] = pd.to_datetime(attendance["Date"], errors="coerce")
    attendance = attendance.dropna(subset=["Timestamp", "Date"])

    if attendance.empty:
        st.warning("No valid attendance records found.")
        return

    week_options = get_week_options(attendance["Date"])
    week_labels = [w[0] for w in week_options]

    st.markdown("### Select Payroll Period")
    c_week1, c_week2 = st.columns(2)

    with c_week1:
        start_week_label = st.selectbox("Start Week", week_labels, index=min(len(week_labels) - 1, 0))

    with c_week2:
        end_week_label = st.selectbox("End Week", week_labels, index=0)

    start_week = next(w for w in week_options if w[0] == start_week_label)
    end_week = next(w for w in week_options if w[0] == end_week_label)

    start_date = min(start_week[1], end_week[1])
    end_date = max(start_week[2], end_week[2])

    selected_label = f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}"

    project_options = ["All Projects"] + sorted(attendance["Project"].dropna().astype(str).unique().tolist())
    selected_project = st.selectbox("Project Filter", project_options)

    filtered = attendance[
        (attendance["Date"].dt.date >= start_date) &
        (attendance["Date"].dt.date <= end_date)
    ].copy()

    if selected_project != "All Projects":
        filtered = filtered[filtered["Project"].astype(str) == selected_project]

    if filtered.empty:
        st.warning("No attendance records found for this selected week/project.")
        return

    results = []
    daily_rows = []

    for name in sorted(filtered["Name"].dropna().astype(str).unique()):
        worker_records = filtered[filtered["Name"].astype(str) == name].sort_values("Timestamp")
        total_normal_hours = 0.0
        total_sunday_hours = 0.0
        total_lunch_hours = 0.0

        for work_date in sorted(worker_records["Date"].dt.date.unique()):
            day_records = worker_records[worker_records["Date"].dt.date == work_date].sort_values("Timestamp")

            check_in = day_records[day_records["Action"] == "Check In"]["Timestamp"]
            lunch_start = day_records[day_records["Action"] == "Lunch Start"]["Timestamp"]
            lunch_end = day_records[day_records["Action"] == "Lunch End"]["Timestamp"]
            check_out = day_records[day_records["Action"] == "Check Out"]["Timestamp"]

            if check_in.empty or check_out.empty:
                daily_rows.append({
                    "Worker": name,
                    "Date": work_date.strftime("%d/%m/%Y"),
                    "Project": ", ".join(sorted(day_records["Project"].dropna().astype(str).unique())),
                    "Check In": "Missing",
                    "Lunch Start": "",
                    "Lunch End": "",
                    "Check Out": "Missing",
                    "Lunch Hours": 0,
                    "Payable Hours": 0,
                    "Note": "Missing check in or check out"
                })
                continue

            start = check_in.iloc[0]
            end = check_out.iloc[-1]
            gross_hours = max((end - start).total_seconds() / 3600, 0)

            lunch_hours = 0.0
            lunch_note = "No lunch recorded"
            lunch_start_text = ""
            lunch_end_text = ""

            if not lunch_start.empty and not lunch_end.empty:
                lunch_s = lunch_start.iloc[0]
                lunch_e = lunch_end.iloc[-1]
                lunch_hours = max((lunch_e - lunch_s).total_seconds() / 3600, 0)
                lunch_start_text = lunch_s.strftime("%I:%M %p")
                lunch_end_text = lunch_e.strftime("%I:%M %p")
                lunch_note = "Lunch deducted"
            elif not lunch_start.empty or not lunch_end.empty:
                lunch_note = "Incomplete lunch record"

            payable_hours = max(gross_hours - lunch_hours, 0)
            total_lunch_hours += lunch_hours

            if work_date.weekday() == 6:
                total_sunday_hours += payable_hours
                pay_note = "Sunday hours"
            else:
                total_normal_hours += payable_hours
                pay_note = "Normal"

            daily_rows.append({
                "Worker": name,
                "Date": work_date.strftime("%d/%m/%Y"),
                "Project": ", ".join(sorted(day_records["Project"].dropna().astype(str).unique())),
                "Check In": start.strftime("%I:%M %p"),
                "Lunch Start": lunch_start_text,
                "Lunch End": lunch_end_text,
                "Check Out": end.strftime("%I:%M %p"),
                "Lunch Hours": round(lunch_hours, 2),
                "Payable Hours": round(payable_hours, 2),
                "Note": f"{pay_note} - {lunch_note}"
            })

        worker_row = workers[workers["Name"].astype(str) == name]
        rate = 35.0
        sunday_rate = 0.0
        bsb = ""
        account = ""
        email = ""

        if not worker_row.empty:
            rate = safe_float(worker_row.iloc[0].get("Rate", 35), 35.0)
            # Sunday Rate starts at 0 in the weekly dashboard.
            # You can manually edit it only for the workers who need Sunday pay.
            sunday_rate = 0.0
            bsb = normalize_bsb(worker_row.iloc[0].get("BSB", ""))
            account = normalize_account(worker_row.iloc[0].get("Account", ""))
            email = clean_text(worker_row.iloc[0].get("Email", ""))

        normal_pay = total_normal_hours * rate
        sunday_pay = total_sunday_hours * sunday_rate
        total_hours = total_normal_hours + total_sunday_hours
        total_pay = normal_pay + sunday_pay

        results.append({
            "Name": name,
            "BSB": bsb,
            "Account": account,
            "Email": email,
            "Normal Hours": round(total_normal_hours, 2),
            "Sunday Hours": round(total_sunday_hours, 2),
            "Total Hours": round(total_hours, 2),
            "Lunch Deducted": round(total_lunch_hours, 2),
            "Rate": round(rate, 2),
            "Sunday Rate": round(sunday_rate, 2),
            "Total Pay": round(total_pay, 2)
        })

    report = pd.DataFrame(results)
    daily_report = pd.DataFrame(daily_rows)

    st.markdown("### Worker Selection")
    report_type = st.radio(
        "Report Type",
        ["All workers", "Selected workers only"],
        horizontal=True
    )

    if report_type == "Selected workers only":
        selected_workers = st.multiselect(
            "Select workers to include",
            report["Name"].tolist(),
            default=report["Name"].tolist()
        )
        report = report[report["Name"].isin(selected_workers)].copy()
        daily_report = daily_report[daily_report["Worker"].isin(selected_workers)].copy()

    if report.empty:
        st.warning("No workers selected for this report.")
        return

    if "Adjustment AUD" not in report.columns:
        report["Adjustment AUD"] = 0.0

    st.markdown("### Weekly Payroll Summary")
    st.caption("You can manually edit Rate, Sunday Rate and Adjustment AUD before downloading or sending the report.")

    editable_report = st.data_editor(
        report,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Name": st.column_config.TextColumn("Name", disabled=True),
            "BSB": st.column_config.TextColumn("BSB", disabled=True),
            "Account": st.column_config.TextColumn("Account", disabled=True),
            "Email": st.column_config.TextColumn("Email", disabled=True),
            "Normal Hours": st.column_config.NumberColumn("Normal Hours", disabled=True, format="%.2f"),
            "Sunday Hours": st.column_config.NumberColumn("Sunday Hours", disabled=True, format="%.2f"),
            "Total Hours": st.column_config.NumberColumn("Total Hours", disabled=True, format="%.2f"),
            "Lunch Deducted": st.column_config.NumberColumn("Lunch Deducted", disabled=True, format="%.2f"),
            "Rate": st.column_config.NumberColumn("Rate", min_value=0.0, step=1.0, format="%.2f"),
            "Sunday Rate": st.column_config.NumberColumn("Sunday Rate", min_value=0.0, step=1.0, format="%.2f"),
            "Adjustment AUD": st.column_config.NumberColumn("Adjustment AUD", step=10.0, format="%.2f"),
            "Total Pay": st.column_config.NumberColumn("Total Pay", disabled=True, format="%.2f"),
        },
        key="weekly_payroll_editor"
    )

    for col in ["Normal Hours", "Sunday Hours", "Rate", "Sunday Rate", "Adjustment AUD"]:
        editable_report[col] = pd.to_numeric(editable_report[col], errors="coerce").fillna(0)

    editable_report["Total Hours"] = editable_report["Normal Hours"] + editable_report["Sunday Hours"]
    editable_report["Total Pay"] = (
        editable_report["Normal Hours"] * editable_report["Rate"]
        + editable_report["Sunday Hours"] * editable_report["Sunday Rate"]
        + editable_report["Adjustment AUD"]
    ).round(2)

    total_workers = len(editable_report)
    total_hours = editable_report["Total Hours"].sum() if not editable_report.empty else 0
    total_pay = editable_report["Total Pay"].sum() if not editable_report.empty else 0

    m1, m2, m3 = st.columns(3)
    m1.metric("Workers", total_workers)
    m2.metric("Total Payable Hours", f"{total_hours:.2f}")
    m3.metric("Total Payroll", f"AUD {total_pay:,.2f}")

    st.markdown("### Final Report Preview")
    preview_report = editable_report.copy()
    if "Email" in preview_report.columns:
        preview_report = preview_report.drop(columns=["Email"])
    st.dataframe(preview_report, use_container_width=True, hide_index=True)

    st.markdown("### Daily Attendance Breakdown")
    st.dataframe(daily_report, use_container_width=True, hide_index=True)

    payroll_filename = f"payroll_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"

    report_for_aaron = editable_report.copy()
    if "Email" in report_for_aaron.columns:
        report_for_aaron = report_for_aaron.drop(columns=["Email"])

    excel_bytes = make_payroll_excel(report_for_aaron, daily_report, selected_label)

    st.markdown("### Download or Email Payroll Report")

    st.download_button(
        "Download Payroll Report Excel",
        data=excel_bytes,
        file_name=payroll_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.markdown("### Send Report to Aaron")
    default_receiver = st.secrets.get("EMAIL_RECEIVER", "")
    email_to = st.text_input("Aaron email", value=default_receiver)
    email_subject = st.text_input("Aaron email subject", value=f"Weekly Payroll Report - {selected_label}")

    aaron_email_preview = build_payroll_email_html(
        report_for_aaron,
        daily_report,
        selected_label,
        total_hours,
        total_pay,
        custom_message=""
    )

    with st.expander("Preview Aaron email", expanded=True):
        st.markdown(aaron_email_preview, unsafe_allow_html=True)

    if st.button("Send Payroll Report to Aaron", use_container_width=True):
        if not email_to.strip():
            st.warning("Please enter Aaron's email address.")
        else:
            try:
                send_payroll_email(
                    to_email=email_to.strip(),
                    subject=email_subject.strip(),
                    html_body=aaron_email_preview,
                    attachment_bytes=excel_bytes,
                    attachment_filename=payroll_filename
                )
                st.success("Payroll report sent to Aaron successfully.")
            except Exception as e:
                st.error("Could not send payroll email to Aaron.")
                st.exception(e)

    st.markdown("### Send Individual Worker Email")
    worker_email_options = editable_report["Name"].dropna().astype(str).tolist()
    selected_worker_email = st.selectbox("Select worker", worker_email_options)

    worker_summary = editable_report[editable_report["Name"].astype(str) == selected_worker_email].copy()
    worker_daily = daily_report[daily_report["Worker"].astype(str) == selected_worker_email].copy()

    default_worker_email = ""
    if not worker_summary.empty and "Email" in worker_summary.columns:
        default_worker_email = str(worker_summary.iloc[0].get("Email", "")).strip()

    worker_to = st.text_input("Worker email", value=default_worker_email)
    worker_subject = st.text_input("Worker email subject", value=f"Your Attendance Summary - {selected_label}")

    worker_email_preview = build_worker_email_html(
        selected_worker_email,
        worker_summary,
        worker_daily,
        selected_label,
        custom_message=""
    )

    with st.expander("Preview worker email", expanded=True):
        st.markdown(worker_email_preview, unsafe_allow_html=True)

    if st.button("Send Worker Attendance Email", use_container_width=True):
        if not worker_to.strip():
            st.warning("Please enter the worker email address.")
        else:
            try:
                send_payroll_email(
                    to_email=worker_to.strip(),
                    subject=worker_subject.strip(),
                    html_body=worker_email_preview,
                    attachment_bytes=None,
                    attachment_filename=None
                )
                st.success(f"Attendance email sent to {selected_worker_email}.")
            except Exception as e:
                st.error("Could not send worker email.")
                st.exception(e)


def projects_page():
    st.title("🏗️ Project Settings")
    st.caption("Add projects here so workers can choose them from the QR attendance form.")

    try:
        spreadsheet = connect_spreadsheet()
        projects_ws = get_or_create_worksheet(spreadsheet, PROJECTS_SHEET_NAME, PROJECTS_HEADERS)
    except Exception as e:
        st.error("Could not connect to Google Sheets.")
        st.exception(e)
        return

    projects = pd.DataFrame(get_sheet_records_cached(PROJECTS_SHEET_NAME, tuple(PROJECTS_HEADERS)))
    if projects.empty:
        projects = pd.DataFrame(columns=PROJECTS_HEADERS)

    st.markdown("### Current Projects")
    st.dataframe(projects, use_container_width=True, hide_index=True)

    st.markdown("### Add New Project")
    c1, c2 = st.columns([3, 1])
    with c1:
        new_project = st.text_input("Project Name")
    with c2:
        active = st.selectbox("Active", ["Yes", "No"])

    if st.button("➕ Add Project", use_container_width=True):
        if not new_project.strip():
            st.warning("Please enter a project name.")
            return

        existing = []
        if not projects.empty and "Project" in projects.columns:
            existing = projects["Project"].dropna().astype(str).str.strip().str.lower().tolist()

        if new_project.strip().lower() in existing:
            st.info("This project already exists.")
            return

        projects_ws.append_row([new_project.strip(), active], value_input_option="RAW")
        st.success(f"Project added: {new_project.strip()}")
        st.rerun()

# =========================================================
# SIDEBAR NAVIGATION / QR WORKER MODE
# =========================================================

try:
    url_mode = str(st.query_params.get("mode", "")).lower().strip()
    url_page = str(st.query_params.get("page", "")).lower().strip()
except Exception:
    url_mode = ""
    url_page = ""

# QR worker mode: this opens only the attendance screen and hides the ERP menu.
# Use this URL for the QR:
# https://aarons-pricing.streamlit.app/?mode=attendance
if url_mode in ["attendance", "worker", "qr", "checkin"]:
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"] {display: none;}
        [data-testid="collapsedControl"] {display: none;}
        .block-container {padding-top: 2rem; max-width: 950px;}
        </style>
        """,
        unsafe_allow_html=True
    )
    attendance_page()
    st.stop()

navigation_pages = [
    "New Costing",
    "Saved Costings",
    "Manager View",
    "Edit Costing",
    "Attendance Register",
    "Payroll Report",
    "Project Settings"
]

default_page_index = 0
if url_page in ["attendance", "checkin", "qr"]:
    default_page_index = navigation_pages.index("Attendance Register")
elif url_page in ["payroll"]:
    default_page_index = navigation_pages.index("Payroll Report")
elif url_page in ["projects", "project-settings"]:
    default_page_index = navigation_pages.index("Project Settings")

page = st.sidebar.radio(
    "Navigation",
    navigation_pages,
    index=default_page_index
)

st.title("🧾 Aaron's Pricing Calculator")
st.caption("Demolition, strip-out, flooring and rubbish removal cost calculator.")

if page == "Attendance Register":
    attendance_page()
    st.stop()

if page == "Payroll Report":
    payroll_page()
    st.stop()

if page == "Project Settings":
    projects_page()
    st.stop()

# =========================================================
# HELPERS
# =========================================================

def money(value):
    try:
        return f"AUD {float(value):,.2f}"
    except Exception:
        return str(value)


def parse_number(value):
    try:
        cleaned = (
            str(value)
            .replace("AUD", "")
            .replace(",", "")
            .replace("%", "")
            .strip()
        )
        if cleaned == "":
            return 0.0
        return float(cleaned)
    except Exception:
        return 0.0


def sale_price_from_margin(cost, margin_pct):
    if margin_pct <= 0 or margin_pct >= 100:
        return cost
    return cost / (1 - margin_pct / 100)


def get_room_options(project_type):
    room_options_map = {
        "Residential": [
            "Kitchen", "Bathroom", "Laundry", "Bedroom", "Living Room",
            "Dining Room", "Hallway", "Garage", "Backyard", "Other"
        ],
        "Commercial": [
            "Tenancy Area", "Office", "Reception", "Meeting Room", "Kitchenette",
            "Bathroom", "Storage", "Shopfront", "Common Area", "Other"
        ],
        "Industrial": [
            "Warehouse Floor", "Racking Area", "Loading Bay", "Office Area",
            "Mezzanine", "Yard", "Concrete Area", "Storage Area", "Waste Area", "Other"
        ],
        "Shopping Centre": [
            "Tenancy Area", "Food Court", "Shopfront", "Kitchen", "Coolroom",
            "Common Area", "Back of House", "Storage", "Other"
        ],
        "Warehouse": [
            "Warehouse Floor", "Racking Area", "Office Area", "Loading Bay",
            "Mezzanine", "Yard", "Waste Area", "Other"
        ],
        "Other": [
            "Area 1", "Area 2", "Room", "External Area", "Internal Area", "Other"
        ],
    }
    return room_options_map.get(project_type, room_options_map["Other"])


def get_item_productivity(item_name, rates_df):
    try:
        match = rates_df[rates_df["item"].astype(str) == str(item_name)]
        if match.empty:
            return 0.0
        return float(match.iloc[0]["productivity_per_day"])
    except Exception:
        return 0.0


def recalc_worker_days_from_saved_rows(df, rates_df):
    total = 0.0
    for _, row in df.iterrows():
        qty = parse_number(row.get("quantity", 0))
        productivity = get_item_productivity(row.get("item", ""), rates_df)
        if productivity > 0:
            total += qty / productivity
    return total


def check_headers_or_warn(df):
    if df.empty:
        return True

    missing = [h for h in EXPECTED_HEADERS if h not in df.columns]
    extra = [c for c in df.columns if c not in EXPECTED_HEADERS]

    if missing or extra:
        st.warning("Costing_Records headers do not exactly match the expected app headers.")
        with st.expander("Expected headers"):
            st.code("\t".join(EXPECTED_HEADERS))
        if missing:
            st.error(f"Missing headers: {missing}")
        if extra:
            st.info(f"Extra headers detected: {extra}")
        return False

    return True

# =========================================================
# GOOGLE SHEETS CONNECTION
# =========================================================

def connect_google_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )

    client = gspread.authorize(creds)
    spreadsheet = client.open(SPREADSHEET_NAME)

    try:
        ws = spreadsheet.worksheet(COSTING_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(
            title=COSTING_SHEET_NAME,
            rows="5000",
            cols="50"
        )
        ws.append_row(EXPECTED_HEADERS, value_input_option="RAW")

    return ws


def load_saved_costings():
    ws = connect_google_sheet()
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=EXPECTED_HEADERS)

    df = pd.DataFrame(records)

    for h in EXPECTED_HEADERS:
        if h not in df.columns:
            df[h] = ""

    return df


def update_costing_records(quote_id, updated_df):
    ws = connect_google_sheet()
    all_values = ws.get_all_values()

    if not all_values:
        ws.append_row(EXPECTED_HEADERS, value_input_option="RAW")
        all_values = [EXPECTED_HEADERS]

    headers = all_values[0]

    if headers != EXPECTED_HEADERS:
        raise ValueError("Costing_Records headers do not exactly match the expected app headers.")

    quote_col_index = headers.index("quote_id")
    rows_to_keep = [headers]

    for row in all_values[1:]:
        row = row + [""] * (len(headers) - len(row))
        if row[quote_col_index] != quote_id:
            rows_to_keep.append(row[:len(headers)])

    ws.clear()
    ws.update("A1", rows_to_keep, value_input_option="RAW")

    if not updated_df.empty:
        updated_df = updated_df[EXPECTED_HEADERS].copy()
        ws.append_rows(updated_df.values.tolist(), value_input_option="RAW")

# =========================================================
# LOAD RATES
# =========================================================

@st.cache_data(ttl=300)
def load_rates(url):
    df = pd.read_csv(url)
    df.columns = [c.strip().lower() for c in df.columns]

    required = [
        "item", "category", "unit", "market_min", "market_recommended",
        "market_high", "productivity_per_day", "notes"
    ]

    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"Missing columns in Rates sheet: {missing}")
        st.stop()

    for c in ["market_min", "market_recommended", "market_high", "productivity_per_day"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    for c in ["item", "category", "unit", "notes"]:
        df[c] = df[c].astype(str).str.strip()

    return df

rates_df = load_rates(SHEET_CSV_URL)

# =========================================================
# QUOTE SELECTOR HELPERS
# =========================================================

def build_quote_options(saved_df):
    if saved_df.empty:
        return pd.DataFrame(columns=["quote_id", "project_address", "client_name", "display"])

    quote_options = (
        saved_df.groupby(
            ["quote_id", "project_address", "client_name"],
            as_index=False
        )
        .size()
    )

    quote_options["display"] = (
        quote_options["project_address"].astype(str)
        + " - "
        + quote_options["client_name"].astype(str)
    )

    quote_options = quote_options.sort_values(by="display", ascending=True)
    return quote_options

# =========================================================
# SAVED COSTINGS PAGE
# =========================================================

if page == "Saved Costings":

    st.title("📂 Saved Costings")

    saved_df = load_saved_costings()

    if saved_df.empty:
        st.warning("No saved costings found.")
        st.stop()

    check_headers_or_warn(saved_df)

    c1, c2, c3 = st.columns(3)

    with c1:
        client_filter = st.text_input("Search Client")

    with c2:
        project_filter = st.text_input("Search Project Address")

    with c3:
        quote_filter = st.text_input("Search Quote ID")

    if client_filter:
        saved_df = saved_df[
            saved_df["client_name"].astype(str).str.contains(client_filter, case=False, na=False)
        ]

    if project_filter:
        saved_df = saved_df[
            saved_df["project_address"].astype(str).str.contains(project_filter, case=False, na=False)
        ]

    if quote_filter:
        saved_df = saved_df[
            saved_df["quote_id"].astype(str).str.contains(quote_filter, case=False, na=False)
        ]

    summary_df = saved_df.groupby(
        ["quote_id", "date", "client_name", "project_address"],
        as_index=False
    ).agg({"final_total": "max"})

    st.markdown("## Costing Records")
    st.dataframe(summary_df, use_container_width=True)

    quote_options = build_quote_options(saved_df)

    if quote_options.empty:
        st.warning("No matching costing found.")
        st.stop()

    selected_display = st.selectbox(
        "Select Quote",
        quote_options["display"].tolist()
    )

    selected_quote = quote_options.loc[
        quote_options["display"] == selected_display,
        "quote_id"
    ].iloc[0]

    selected_df = saved_df[saved_df["quote_id"] == selected_quote]

    if not selected_df.empty:
        st.markdown("### Full Costing Breakdown")
        st.dataframe(selected_df, use_container_width=True)

    st.stop()

# =========================================================
# MANAGER VIEW - COMPACT AND FIXED DURATION
# =========================================================

if page == "Manager View":

    st.title("📊 Manager Costing View")

    saved_df = load_saved_costings()

    if saved_df.empty:
        st.warning("No saved costings found.")
        st.stop()

    check_headers_or_warn(saved_df)

    quote_options = build_quote_options(saved_df)

    if quote_options.empty:
        st.warning("No saved costings found.")
        st.stop()

    selected_display = st.selectbox(
        "Select Quote",
        quote_options["display"].tolist()
    )

    selected_quote = quote_options.loc[
        quote_options["display"] == selected_display,
        "quote_id"
    ].iloc[0]

    selected_df = saved_df[saved_df["quote_id"] == selected_quote].copy()

    if selected_df.empty:
        st.warning("No records found for this quote.")
        st.stop()

    first_row = selected_df.iloc[0]

    st.markdown(
        """
        <style>
        div[data-testid="stMetricValue"] {
            font-size: 24px;
        }
        div[data-testid="stMetricLabel"] {
            font-size: 13px;
        }
        .block-container {
            padding-top: 1.5rem;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown("### Project Information")
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Client", str(first_row["client_name"]))
    p2.metric("Project Type", str(first_row["project_type"]))
    p3.metric("Working Hours", str(first_row["working_hours"]))
    p4.metric("Quote ID", str(first_row["quote_id"]))

    st.caption(str(first_row["project_address"]))

    st.divider()

    st.markdown("### Financial Summary")

    direct_cost = str(first_row["direct_cost"])
    recommended_price = str(first_row["recommended_price"])
    gst_total = str(first_row["gst_total"])
    final_total = str(first_row["final_total"])
    labour_cost = str(first_row["labour_cost"])
    supervisor_cost = str(first_row["supervisor_cost"])
    equipment_cost = str(first_row["equipment_cost"])
    disposal_cost = str(first_row["disposal_cost"])
    truck_cost = str(first_row["truck_cost"])
    difficulty_allowance = str(first_row["difficulty_allowance"])

    f1, f2, f3, f4 = st.columns(4)
    f1.metric("Direct Cost", direct_cost)
    f2.metric("Recommended", recommended_price)
    f3.metric("GST", gst_total)
    f4.metric("Final Total", final_total)

    f5, f6, f7, f8 = st.columns(4)
    f5.metric("Labour", labour_cost)
    f6.metric("Supervisor", supervisor_cost)
    f7.metric("Equipment", equipment_cost)
    f8.metric("Disposal / Truck", f"{disposal_cost} / {truck_cost}")

    st.caption(f"Difficulty Allowance: {difficulty_allowance}")

    st.divider()

    st.markdown("### Program Summary")

    total_worker_days = recalc_worker_days_from_saved_rows(selected_df, rates_df)
    labourers = 2
    estimated_duration = total_worker_days / labourers if labourers > 0 else 0
    rounded_duration = ceil(estimated_duration) if estimated_duration > 0 else 0

    d1, d2, d3 = st.columns(3)
    d1.metric("Estimated Worker Days", f"{total_worker_days:.2f}")
    d2.metric("Labourers", str(labourers))
    d3.metric("Estimated Duration", f"{estimated_duration:.2f} days ({rounded_duration} rounded)")

    st.divider()

    st.markdown("### Work Breakdown")

    manager_df = selected_df[
        ["room_area", "item", "quantity", "rate", "item_total", "notes"]
    ].copy()

    manager_df.columns = [
        "Area", "Work Item", "Quantity", "Rate", "Total", "Notes"
    ]

    st.dataframe(manager_df, use_container_width=True, hide_index=True)

    st.stop()

# =========================================================
# EDIT COSTING - BASIC EDITOR WITH DROPDOWNS
# =========================================================

if page == "Edit Costing":

    st.title("✏️ Edit Costing")

    saved_df = load_saved_costings()

    if saved_df.empty:
        st.warning("No saved costings found.")
        st.stop()

    if not check_headers_or_warn(saved_df):
        st.stop()

    quote_options = build_quote_options(saved_df)

    selected_display = st.selectbox(
        "Select Quote to Edit",
        quote_options["display"].tolist()
    )

    selected_quote = quote_options.loc[
        quote_options["display"] == selected_display,
        "quote_id"
    ].iloc[0]

    selected_df = saved_df[saved_df["quote_id"] == selected_quote].copy()

    if selected_df.empty:
        st.warning("No records found for this quote.")
        st.stop()

    first_row = selected_df.iloc[0]

    st.markdown("### Project Details")

    c1, c2 = st.columns(2)

    with c1:
        edit_client = st.text_input("Client Name", value=str(first_row["client_name"]))
        edit_address = st.text_input("Project Address", value=str(first_row["project_address"]))

    with c2:
        project_type_options = ["Residential", "Commercial", "Industrial", "Shopping Centre", "Warehouse", "Other"]
        current_project_type = str(first_row["project_type"])
        if current_project_type not in project_type_options:
            project_type_options.append(current_project_type)

        edit_project_type = st.selectbox(
            "Project Type",
            project_type_options,
            index=project_type_options.index(current_project_type)
        )

        working_hours_options = ["Business Hours", "After Hours", "Weekend", "Day & Night", "To be confirmed"]
        current_working_hours = str(first_row["working_hours"])
        if current_working_hours not in working_hours_options:
            working_hours_options.append(current_working_hours)

        edit_working_hours = st.selectbox(
            "Working Hours",
            working_hours_options,
            index=working_hours_options.index(current_working_hours)
        )

    st.markdown("### Edit Work Items")
    st.caption("You can edit area, category, item, measurement method, quantity, rate and notes. Totals recalculate below.")

    edit_cols = [
        "room_area",
        "item",
        "category",
        "unit",
        "measurement_method",
        "quantity",
        "rate",
        "notes"
    ]

    editable_df = selected_df[edit_cols].copy()
    editable_df = editable_df.astype(str)

    item_options = sorted(rates_df["item"].dropna().astype(str).unique().tolist())
    category_options = sorted(rates_df["category"].dropna().astype(str).unique().tolist())
    unit_options = sorted(rates_df["unit"].dropna().astype(str).unique().tolist())
    room_options = get_room_options(edit_project_type)
    current_rooms = editable_df["room_area"].dropna().astype(str).unique().tolist()
    for room in current_rooms:
        if room not in room_options:
            room_options.append(room)

    measurement_options = [
        "Manual Quantity",
        "Wall Calculator (Total Length x Height)",
        "Area Calculator (Length x Width)"
    ]

    edited_df = st.data_editor(
        editable_df,
        use_container_width=True,
        num_rows="dynamic",
        key="edit_costing_editor",
        column_config={
            "room_area": st.column_config.SelectboxColumn(
                "Area",
                options=room_options,
                required=True
            ),
            "category": st.column_config.SelectboxColumn(
                "Category",
                options=category_options,
                required=True
            ),
            "item": st.column_config.SelectboxColumn(
                "Work Item",
                options=item_options,
                required=True
            ),
            "unit": st.column_config.SelectboxColumn(
                "Unit",
                options=unit_options,
                required=True
            ),
            "measurement_method": st.column_config.SelectboxColumn(
                "Measurement Method",
                options=measurement_options,
                required=True
            ),
            "quantity": st.column_config.TextColumn("Quantity"),
            "rate": st.column_config.TextColumn("Rate"),
            "notes": st.column_config.TextColumn("Notes"),
        }
    )

    st.markdown("### Costs & Settings")

    r1, r2, r3, r4 = st.columns(4)

    with r1:
        edit_equipment_cost = st.text_input("Equipment Cost", value=str(first_row["equipment_cost"]))

    with r2:
        edit_disposal_cost = st.text_input("Disposal / Tip Fees", value=str(first_row["disposal_cost"]))

    with r3:
        edit_truck_cost = st.text_input("Truck / Transport Cost", value=str(first_row["truck_cost"]))

    with r4:
        edit_other_cost = st.text_input("Other Cost", value=str(first_row["other_cost"]))

    r5, r6, r7, r8 = st.columns(4)

    with r5:
        edit_labour_cost = st.text_input("Labour Cost", value=str(first_row["labour_cost"]))

    with r6:
        edit_supervisor_cost = st.text_input("Supervisor Cost", value=str(first_row["supervisor_cost"]))

    with r7:
        edit_difficulty_percent = st.text_input("Difficulty %", value=str(first_row["difficulty_percent"]))

    with r8:
        edit_gst_percent = st.number_input("GST %", min_value=0.0, max_value=20.0, value=10.0, step=0.5)

    recalculated_rows = []

    for _, row in edited_df.iterrows():
        item_name = str(row.get("item", "")).strip()
        qty = parse_number(row.get("quantity", 0))
        rate = parse_number(row.get("rate", 0))
        item_total = qty * rate

        rate_match = rates_df[rates_df["item"].astype(str) == item_name]

        if not rate_match.empty:
            rate_row = rate_match.iloc[0]
            category = str(rate_row["category"])
            unit = str(rate_row["unit"])
            productivity = float(rate_row["productivity_per_day"])
        else:
            category = str(row.get("category", ""))
            unit = str(row.get("unit", ""))
            productivity = 0.0

        worker_days = qty / productivity if productivity > 0 else 0.0

        recalculated_rows.append({
            "room_area": str(row.get("room_area", "")),
            "item": item_name,
            "category": category,
            "unit": unit,
            "measurement_method": str(row.get("measurement_method", "")),
            "quantity": qty,
            "rate": rate,
            "item_total": item_total,
            "worker_days": worker_days,
            "notes": str(row.get("notes", "")),
        })

    recalculated_df = pd.DataFrame(recalculated_rows)

    base_total = recalculated_df["item_total"].sum() if not recalculated_df.empty else 0.0

    equipment_num = parse_number(edit_equipment_cost)
    disposal_num = parse_number(edit_disposal_cost)
    truck_num = parse_number(edit_truck_cost)
    other_num = parse_number(edit_other_cost)
    labour_num = parse_number(edit_labour_cost)
    supervisor_num = parse_number(edit_supervisor_cost)
    difficulty_num = parse_number(edit_difficulty_percent)

    direct_cost = (
        base_total
        + equipment_num
        + disposal_num
        + truck_num
        + other_num
        + labour_num
        + supervisor_num
    )

    difficulty_allowance = direct_cost * (difficulty_num / 100)
    cost_with_difficulty = direct_cost + difficulty_allowance
    recommended_price = sale_price_from_margin(cost_with_difficulty, 30.0)
    gst_total = recommended_price * (edit_gst_percent / 100)
    final_total = recommended_price + gst_total

    total_worker_days = recalculated_df["worker_days"].sum() if not recalculated_df.empty else 0.0
    labourers = 2
    estimated_duration = total_worker_days / labourers if labourers > 0 else 0
    rounded_duration = ceil(estimated_duration) if estimated_duration > 0 else 0

    st.markdown("### Updated Summary")

    s1, s2, s3 = st.columns(3)
    s1.metric("Direct Cost", money(direct_cost))
    s2.metric("Recommended Price", money(recommended_price))
    s3.metric("Final Total Inc GST", money(final_total))

    s4, s5, s6 = st.columns(3)
    s4.metric("Estimated Worker Days", f"{total_worker_days:.2f}")
    s5.metric("Labourers", str(labourers))
    s6.metric("Estimated Duration", f"{estimated_duration:.2f} days ({rounded_duration} rounded)")

    st.markdown("### Recalculated Work Items")

    preview_df = recalculated_df.copy()
    if not preview_df.empty:
        preview_df["rate"] = preview_df["rate"].apply(money)
        preview_df["item_total"] = preview_df["item_total"].apply(money)
        preview_df["worker_days"] = preview_df["worker_days"].map(lambda x: f"{float(x):.2f}")
        st.dataframe(preview_df, use_container_width=True, hide_index=True)

    st.markdown("### Save Changes")

    if st.button("💾 Update Costing"):

        now = melbourne_now().strftime("%d/%m/%Y %H:%M")
        updated_rows = []

        for _, row in recalculated_df.iterrows():
            updated_rows.append({
                "quote_id": selected_quote,
                "date": now,
                "client_name": edit_client,
                "project_address": edit_address,
                "project_type": edit_project_type,
                "working_hours": edit_working_hours,
                "floor_level": str(first_row["floor_level"]),
                "room_area": str(row["room_area"]),
                "item": str(row["item"]),
                "category": str(row["category"]),
                "unit": str(row["unit"]),
                "measurement_method": str(row["measurement_method"]),
                "quantity": round(float(row["quantity"]), 2),
                "rate": money(row["rate"]),
                "item_total": money(row["item_total"]),
                "worker_days": round(float(row["worker_days"]), 2),
                "equipment_cost": money(equipment_num),
                "disposal_cost": money(disposal_num),
                "truck_cost": money(truck_num),
                "other_cost": money(other_num),
                "difficulty_percent": f"{difficulty_num:.1f}%",
                "difficulty_allowance": money(difficulty_allowance),
                "labour_cost": money(labour_num),
                "supervisor_cost": money(supervisor_num),
                "direct_cost": money(direct_cost),
                "recommended_price": money(recommended_price),
                "gst_total": money(gst_total),
                "final_total": money(final_total),
                "notes": str(row["notes"])
            })

        final_df = pd.DataFrame(updated_rows)
        final_df = final_df[EXPECTED_HEADERS]

        try:
            update_costing_records(selected_quote, final_df)
            st.success(f"✅ Quote updated successfully: {selected_quote}")
        except Exception as e:
            st.error("Could not update costing.")
            st.exception(e)

    st.stop()

# =========================================================
# SESSION STATE FOR NEW COSTING
# =========================================================

if "room_count" not in st.session_state:
    st.session_state.room_count = 1

if "room_items" not in st.session_state:
    st.session_state.room_items = {0: 1}

if "reset_counter" not in st.session_state:
    st.session_state.reset_counter = 0


def add_room():
    idx = st.session_state.room_count
    st.session_state.room_count += 1
    st.session_state.room_items[idx] = 1


def remove_room():
    if st.session_state.room_count > 1:
        idx = st.session_state.room_count - 1
        st.session_state.room_items.pop(idx, None)
        st.session_state.room_count -= 1


def add_work_item(room_idx):
    st.session_state.room_items[room_idx] += 1


def remove_work_item(room_idx):
    if st.session_state.room_items[room_idx] > 1:
        st.session_state.room_items[room_idx] -= 1


def reset_all():
    st.session_state.room_count = 1
    st.session_state.room_items = {0: 1}
    st.session_state.reset_counter += 1

# =========================================================
# NEW COSTING PAGE
# =========================================================

st.subheader("1. Project Details")

c1, c2, c3 = st.columns(3)

with c1:
    client_name = st.text_input("Client Name")
    project_address = st.text_input("Project Address")

with c2:
    project_type = st.selectbox(
        "Project Type",
        ["Residential", "Commercial", "Industrial", "Shopping Centre", "Warehouse", "Other"]
    )
    working_hours = st.selectbox(
        "Working Hours",
        ["Business Hours", "After Hours", "Weekend", "Day & Night", "To be confirmed"]
    )

with c3:
    floor_level = st.number_input("Floor Level", min_value=0, value=0, step=1)
    gst_rate = st.number_input("GST %", min_value=0.0, max_value=20.0, value=10.0, step=0.5)

st.divider()

room_options = get_room_options(project_type)

st.subheader("2. Rooms / Areas & Work Items")

b1, b2, b3 = st.columns([1, 1, 4])

with b1:
    st.button("➕ Add Room / Area", on_click=add_room, use_container_width=True)

with b2:
    st.button("➖ Remove Last Room", on_click=remove_room, use_container_width=True)

with b3:
    st.button("🔄 Reset All Rooms", on_click=reset_all)

items = []
categories = ["All"] + sorted(rates_df["category"].dropna().unique().tolist())

for room_idx in range(st.session_state.room_count):
    room_key = f"{st.session_state.reset_counter}_{room_idx}"

    st.markdown(f"## Room / Area {room_idx + 1}")

    r1, r2 = st.columns(2)

    with r1:
        selected_room = st.selectbox("Room / Area Type", room_options, key=f"room_type_{room_key}")

    with r2:
        custom_room_name = st.text_input("Custom Name (Optional)", key=f"custom_room_{room_key}")

    room_name = custom_room_name.strip() if custom_room_name.strip() else selected_room

    a1, a2 = st.columns(2)

    with a1:
        st.button(
            f"➕ Add Work Item - {room_name}",
            key=f"add_item_{room_key}",
            on_click=add_work_item,
            args=(room_idx,),
            use_container_width=True
        )

    with a2:
        st.button(
            f"➖ Remove Last Item - {room_name}",
            key=f"remove_item_{room_key}",
            on_click=remove_work_item,
            args=(room_idx,),
            use_container_width=True
        )

    for item_idx in range(st.session_state.room_items[room_idx]):
        item_key = f"{room_key}_{item_idx}"

        st.markdown(f"### {room_name} - Item {item_idx + 1}")

        x1, x2 = st.columns([1.2, 2])

        with x1:
            category_filter = st.selectbox("Category", categories, key=f"category_{item_key}")

        filtered_df = rates_df if category_filter == "All" else rates_df[rates_df["category"] == category_filter]

        with x2:
            selected_item = st.selectbox("Work Type", filtered_df["item"].tolist(), key=f"work_type_{item_key}")

        row = rates_df[rates_df["item"] == selected_item].iloc[0]

        unit = str(row["unit"]).strip()
        unit_lower = unit.lower()
        item_lower = selected_item.lower()

        st.markdown("#### Measurement")

        measurement_options = ["Manual Quantity"]

        if "wall" in item_lower or "partition" in item_lower:
            measurement_options.append("Wall Calculator (Total Length x Height)")

        if unit_lower in ["m2", "sqm", "sq m", "m²"]:
            measurement_options.append("Area Calculator (Length x Width)")

        measurement_method = st.radio(
            "Measurement Method",
            measurement_options,
            horizontal=True,
            key=f"measure_{item_key}"
        )

        if measurement_method == "Wall Calculator (Total Length x Height)":
            w1, w2 = st.columns(2)

            with w1:
                total_wall_length = st.number_input(
                    "Total Wall Length (lm)",
                    min_value=0.0,
                    value=1.0,
                    step=0.5,
                    key=f"wall_length_{item_key}"
                )

            with w2:
                wall_height = st.number_input(
                    "Wall Height (m)",
                    min_value=0.0,
                    value=2.7,
                    step=0.1,
                    key=f"wall_height_{item_key}"
                )

            wall_area = total_wall_length * wall_height

            if unit_lower in ["lm", "linear metre", "linear meter"]:
                quantity = total_wall_length
                measurement_note = f"Wall area: {wall_area:.2f} m². Pricing quantity used: {quantity:.2f} lm."
            else:
                quantity = wall_area
                measurement_note = f"Wall area used for pricing: {quantity:.2f} m²."

            st.info(measurement_note)

        elif measurement_method == "Area Calculator (Length x Width)":
            l1, l2 = st.columns(2)

            with l1:
                length = st.number_input(
                    "Length (m)",
                    min_value=0.0,
                    value=1.0,
                    step=0.5,
                    key=f"length_{item_key}"
                )

            with l2:
                width = st.number_input(
                    "Width (m)",
                    min_value=0.0,
                    value=1.0,
                    step=0.5,
                    key=f"width_{item_key}"
                )

            quantity = length * width
            measurement_note = f"Area used for pricing: {quantity:.2f} m²."
            st.info(measurement_note)

        else:
            quantity = st.number_input(
                f"Manual Quantity ({unit})",
                min_value=0.0,
                value=1.0,
                step=1.0,
                key=f"qty_{item_key}"
            )
            measurement_note = f"Manual quantity used: {quantity:.2f} {unit}."

        p1, p2, p3 = st.columns([1, 1, 2])

        with p1:
            your_rate = st.number_input(
                f"Your Rate per {unit}",
                min_value=0.0,
                value=float(row["market_recommended"]),
                step=5.0,
                key=f"rate_{item_key}"
            )

        item_total = quantity * your_rate

        with p2:
            st.metric("Item Total", money(item_total))

        with p3:
            st.caption(f"Quantity used: {quantity:.2f} {unit}")

        market_min = float(row["market_min"])
        market_rec = float(row["market_recommended"])
        market_high = float(row["market_high"])
        productivity = float(row["productivity_per_day"])

        if your_rate < market_min:
            st.error(f"🚨 Below market minimum: {money(market_min)}/{unit}")
        elif market_min <= your_rate < market_rec:
            st.warning(f"⚠️ Below recommended: {money(market_rec)}/{unit}")
        elif market_rec <= your_rate <= market_high:
            st.success("✅ Within normal market range.")
        else:
            st.info(f"💎 Premium pricing. Market high: {money(market_high)}/{unit}")

        m1, m2, m3, m4 = st.columns(4)
        m1.caption(f"Market Min: {money(market_min)}/{unit}")
        m2.caption(f"Recommended: {money(market_rec)}/{unit}")
        m3.caption(f"Market High: {money(market_high)}/{unit}")
        m4.caption(f"Productivity: {productivity:g} {unit}/worker/day")

        notes = st.text_input("Item Notes", value=str(row["notes"]), key=f"notes_{item_key}")

        worker_days = quantity / productivity if productivity > 0 else 0

        items.append({
            "Room / Area": room_name,
            "Item": selected_item,
            "Category": row["category"],
            "Unit": unit,
            "Measurement Method": measurement_method,
            "Quantity Used": quantity,
            "Your Rate": your_rate,
            "Total": item_total,
            "Worker Days": worker_days,
            "Notes": notes
        })

        st.divider()

items_df = pd.DataFrame(items)

base_total = float(items_df["Total"].sum()) if not items_df.empty else 0
total_worker_days = float(items_df["Worker Days"].sum()) if not items_df.empty else 0

st.subheader("3. Labour & Estimated Duration")

l1, l2, l3, l4 = st.columns(4)

with l1:
    labourers = st.number_input("Number of Labourers", min_value=1, value=2, step=1)

with l2:
    labourer_day_rate = st.number_input("Labourer Cost / Day", min_value=0.0, value=400.0, step=25.0)

with l3:
    supervisor_days = st.number_input("Supervisor Days", min_value=0.0, value=0.0, step=0.5)

with l4:
    supervisor_day_rate = st.number_input("Supervisor Cost / Day", min_value=0.0, value=550.0, step=25.0)

estimated_days = total_worker_days / labourers if labourers > 0 else 0
rounded_days = ceil(estimated_days) if estimated_days > 0 else 0

labour_cost = rounded_days * labourers * labourer_day_rate
supervisor_cost = supervisor_days * supervisor_day_rate

st.info(f"Estimated labour duration: {estimated_days:.2f} days ({rounded_days} rounded day/s)")

st.subheader("4. Extra Costs")

e1, e2, e3, e4 = st.columns(4)

with e1:
    equipment_cost = st.number_input("Equipment Cost", min_value=0.0, value=0.0, step=100.0)

with e2:
    disposal_cost = st.number_input("Disposal / Tip Fees", min_value=0.0, value=0.0, step=100.0)

with e3:
    truck_cost = st.number_input("Truck / Transport Cost", min_value=0.0, value=0.0, step=100.0)

with e4:
    other_cost = st.number_input("Other Cost", min_value=0.0, value=0.0, step=100.0)

st.divider()

st.subheader("5. Difficulty & Access Multipliers")

default_multipliers = {
    "Lift Access": 10.0,
    "Stairs Access": 15.0,
    "CBD Location": 15.0,
    "Shopping Centre": 20.0,
    "After Hours": 25.0,
    "Weekend Works": 30.0,
    "Rooftop Access": 40.0,
    "Manual Handling": 10.0,
    "Restricted Access": 15.0,
    "Heavy Concrete": 25.0,
    "High Dust Environment": 10.0,
    "Live / Occupied Site": 15.0,
}

selected_multiplier_total = 0.0
multiplier_rows = []
cols = st.columns(3)

for idx, (name, default_pct) in enumerate(default_multipliers.items()):
    with cols[idx % 3]:
        checked = st.checkbox(name, value=False, key=f"check_{name}")
        pct = st.number_input(
            f"{name} %",
            min_value=0.0,
            max_value=100.0,
            value=default_pct,
            step=1.0,
            disabled=not checked,
            key=f"pct_{name}"
        )

        if checked:
            selected_multiplier_total += pct
            multiplier_rows.append({"Condition": name, "Percent": pct})

use_custom = st.checkbox("Custom Extra %", value=False)
custom_pct = st.number_input(
    "Custom Difficulty / Risk %",
    min_value=0.0,
    max_value=100.0,
    value=0.0,
    step=1.0,
    disabled=not use_custom
)

if use_custom:
    selected_multiplier_total += custom_pct
    multiplier_rows.append({"Condition": "Custom Extra", "Percent": custom_pct})

st.subheader("6. Margin & Final Price")

g1, g2, g3 = st.columns(3)

with g1:
    minimum_margin = st.number_input("Minimum Margin %", min_value=0.0, max_value=80.0, value=20.0, step=1.0)

with g2:
    recommended_margin = st.number_input("Recommended Margin %", min_value=0.0, max_value=80.0, value=30.0, step=1.0)

with g3:
    premium_margin = st.number_input("Premium Margin %", min_value=0.0, max_value=80.0, value=40.0, step=1.0)

direct_cost = base_total + labour_cost + supervisor_cost + equipment_cost + disposal_cost + truck_cost + other_cost
difficulty_allowance = direct_cost * (selected_multiplier_total / 100)
cost_with_difficulty = direct_cost + difficulty_allowance

minimum_price = sale_price_from_margin(cost_with_difficulty, minimum_margin)
recommended_price = sale_price_from_margin(cost_with_difficulty, recommended_margin)
premium_price = sale_price_from_margin(cost_with_difficulty, premium_margin)

recommended_gst = recommended_price * (gst_rate / 100)
recommended_inc_gst = recommended_price + recommended_gst

st.divider()
st.subheader("7. Pricing Summary")

s1, s2, s3, s4 = st.columns(4)

s1.metric("Base Work Items", money(base_total))
s2.metric("Labour Cost", money(labour_cost))
s3.metric("Difficulty %", f"{selected_multiplier_total:.1f}%")
s4.metric("Difficulty Allowance", money(difficulty_allowance))

s5, s6, s7, s8 = st.columns(4)

s5.metric("Direct Cost", money(direct_cost))
s6.metric("Cost + Difficulty", money(cost_with_difficulty))
s7.metric("Recommended + GST", money(recommended_price))
s8.metric("Total Inc GST", money(recommended_inc_gst))

st.markdown("### Suggested Prices")

q1, q2, q3 = st.columns(3)

q1.success(f"Minimum Safe Price: **{money(minimum_price)} + GST**")
q2.info(f"Recommended Price: **{money(recommended_price)} + GST**")
q3.warning(f"Premium Price: **{money(premium_price)} + GST**")

st.markdown("### Work Item Breakdown")

display_df = items_df.copy()

if not display_df.empty:
    display_df["Your Rate"] = display_df["Your Rate"].apply(money)
    display_df["Total"] = display_df["Total"].apply(money)
    display_df["Quantity Used"] = display_df["Quantity Used"].map(lambda x: f"{float(x):.2f}")
    display_df["Worker Days"] = display_df["Worker Days"].map(lambda x: f"{float(x):.2f}")

    st.dataframe(display_df, use_container_width=True)

if multiplier_rows:
    st.markdown("### Selected Multipliers")
    st.dataframe(pd.DataFrame(multiplier_rows), use_container_width=True)

st.markdown("### Copy Summary")

summary_text = f"""
Project: {client_name or 'TBC'}
Address: {project_address or 'TBC'}
Project Type: {project_type}
Working Hours: {working_hours}
Floor Level: {floor_level}

Base Work Items: {money(base_total)}
Labour Cost: {money(labour_cost)}
Supervisor Cost: {money(supervisor_cost)}
Equipment Cost: {money(equipment_cost)}
Disposal / Tip Fees: {money(disposal_cost)}
Truck / Transport Cost: {money(truck_cost)}
Other Cost: {money(other_cost)}
Difficulty Allowance ({selected_multiplier_total:.1f}%): {money(difficulty_allowance)}

Minimum Safe Price: {money(minimum_price)} + GST
Recommended Price: {money(recommended_price)} + GST
Premium Price: {money(premium_price)} + GST

Recommended GST: {money(recommended_gst)}
Recommended Total Inc GST: {money(recommended_inc_gst)}
Estimated Duration: {estimated_days:.2f} days ({rounded_days} rounded day/s)
"""

st.text_area("Copy Summary", value=summary_text.strip(), height=320)

st.markdown("## Save Costing")

if st.button("💾 Save Costing to Google Sheets"):
    if items_df.empty:
        st.warning("No work items to save.")
    else:
        try:
            ws = connect_google_sheet()

            quote_id = f"Q-{melbourne_now().strftime('%Y%m%d-%H%M%S')}"
            now = melbourne_now().strftime("%d/%m/%Y %H:%M")

            rows = []

            for _, row in items_df.iterrows():
                rows.append([
                    str(quote_id),
                    str(now),
                    str(client_name),
                    str(project_address),
                    str(project_type),
                    str(working_hours),
                    str(floor_level),
                    str(row["Room / Area"]),
                    str(row["Item"]),
                    str(row["Category"]),
                    str(row["Unit"]),
                    str(row["Measurement Method"]),
                    round(float(row["Quantity Used"]), 2),
                    money(row["Your Rate"]),
                    money(row["Total"]),
                    round(float(row["Worker Days"]), 2),
                    money(equipment_cost),
                    money(disposal_cost),
                    money(truck_cost),
                    money(other_cost),
                    f"{float(selected_multiplier_total):.1f}%",
                    money(difficulty_allowance),
                    money(labour_cost),
                    money(supervisor_cost),
                    money(direct_cost),
                    money(recommended_price),
                    money(recommended_gst),
                    money(recommended_inc_gst),
                    str(row["Notes"])
                ])

            ws.append_rows(rows, value_input_option="RAW")

            st.success(f"✅ Costing saved successfully. Quote ID: {quote_id}")

        except Exception as e:
            st.error("Could not save costing to Google Sheets.")
            st.exception(e)
