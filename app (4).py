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
import html as html_lib
import hashlib
import hmac
import time as time_mod
import streamlit.components.v1 as components

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
ATTENDANCE_HEADERS = ["Timestamp", "Date", "Name", "Project", "Action", "Latitude", "Longitude", "Location Link", "Note"]
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
    """Read records from Google Sheets with a short cache to avoid Google API 429 quota errors.
    Values are read as text so BSB / account numbers keep their leading zeros.
    """
    spreadsheet = connect_spreadsheet()
    ws = get_or_create_worksheet(spreadsheet, sheet_name, list(headers_tuple))
    try:
        return ws.get_all_records(numericise_ignore=["all"])
    except TypeError:
        return ws.get_all_records()


def clear_sheet_cache():
    """Clear cached sheet reads after writing new rows (keeps the rates cache)."""
    try:
        get_sheet_records_cached.clear()
    except Exception:
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


def project_picker(projects_ws, default_project=None, key_prefix="att"):
    project_options = get_project_options(projects_ws)

    if project_options:
        index = project_options.index(default_project) if default_project in project_options else 0
        selected = st.selectbox("🏗️ Project / Site", project_options, index=index)
        with st.expander("Project not in the list?"):
            other_project = st.text_input("Type the project name", value="", key=f"{key_prefix}_other_project")
        return other_project.strip() if other_project.strip() else selected

    st.info("No projects found yet. Type the project name below.")
    return st.text_input("🏗️ Project / Site", value="", key=f"{key_prefix}_project_manual").strip()


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



# =========================================================
# UI HELPERS (Attendance / Payroll / Projects)
# =========================================================

INACTIVE_VALUES = ["no", "inactive", "false", "0"]
REGISTER_OPTIONS = ["✅ I'm already registered", "🆕 First time here"]
REVIEW_STATUS = "⚠️ Review"
OK_STATUS = "✅ OK"
EDITED_STATUS = "✏️ Edited"
VOID_SUFFIX = " (void)"          # original rows replaced by a manual correction stay in the sheet with this suffix
MANUAL_SOURCE = "Manual"         # value written in "Location Link" for manual corrections


def esc(value):
    """Escape text before putting it inside HTML."""
    return html_lib.escape(str(value), quote=True)


def ops_html(markup):
    st.markdown(markup, unsafe_allow_html=True)


def ops_spacer(height="1.75rem"):
    """Small vertical gap so a button lines up with inputs in the same row."""
    ops_html(f"<div style='height:{height}'></div>")


def inject_ops_css():
    ops_html("""<style>
.ops-hero{background:linear-gradient(135deg,#111827 0%,#374151 100%);color:#ffffff;border-radius:14px;padding:18px 22px;margin:0 0 16px 0;}
.ops-hero-title{font-size:1.6rem;font-weight:700;line-height:1.2;color:#ffffff;}
.ops-hero-sub{font-size:0.92rem;color:#d1d5db;margin-top:4px;}
.ops-pill{display:inline-block;padding:4px 12px;border-radius:999px;font-size:0.85rem;font-weight:600;margin:2px 8px 8px 0;}
.ops-pill.ok{background:#dcfce7;color:#166534;}
.ops-pill.warn{background:#fef3c7;color:#92400e;}
.ops-pill.bad{background:#fee2e2;color:#991b1b;}
.ops-link{font-size:0.85rem;}
.ops-steps{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:4px 0 14px 0;}
@media (max-width:640px){.ops-steps{grid-template-columns:repeat(2,1fr);}}
.ops-step{border:1px solid rgba(128,128,128,0.30);border-radius:10px;padding:10px 12px;}
.ops-step.done{border-color:rgba(34,197,94,0.60);background:rgba(34,197,94,0.12);}
.ops-step.next{border-color:rgba(245,158,11,0.80);background:rgba(245,158,11,0.12);}
.ops-step-label{font-size:0.72rem;text-transform:uppercase;letter-spacing:0.04em;opacity:0.70;}
.ops-step-value{font-size:1.05rem;font-weight:700;margin-top:2px;}
.ops-success{border:1px solid rgba(34,197,94,0.60);background:rgba(34,197,94,0.14);border-radius:12px;padding:14px 16px;margin:0 0 14px 0;}
.ops-success-title{font-size:1.15rem;font-weight:700;}
div[data-testid="stMetric"]{border:1px solid rgba(128,128,128,0.25);border-radius:12px;padding:12px 16px;background:rgba(128,128,128,0.05);}
.stButton>button,.stFormSubmitButton>button,.stDownloadButton>button{border-radius:10px;min-height:2.8rem;font-weight:600;}
</style>""")


def ops_hero(title, subtitle=""):
    sub = f"<div class='ops-hero-sub'>{esc(subtitle)}</div>" if subtitle else ""
    ops_html(f"<div class='ops-hero'><div class='ops-hero-title'>{esc(title)}</div>{sub}</div>")


def format_clock(ts):
    return ts.strftime("%I:%M %p").lstrip("0")


def parse_sheet_timestamp(value):
    text = clean_text(value)
    if not text:
        return None
    try:
        return datetime.strptime(text, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        parsed = pd.to_datetime(text, errors="coerce")
        return None if pd.isna(parsed) else parsed.to_pydatetime()


def todays_attendance(records, name, today_str):
    """Return ({action: 'h:mm AM'}, last_project) for one worker today.
    Uses the same rule as payroll: first Check In / Lunch Start, last Lunch End / Check Out.
    """
    rows = []
    for record in records:
        if clean_text(record.get("Name", "")) != name:
            continue
        action = clean_text(record.get("Action", ""))
        if action not in ATTENDANCE_ACTIONS:      # skips voided rows
            continue
        ts = parse_sheet_timestamp(record.get("Timestamp", ""))
        if ts is None:
            continue
        record_date = clean_text(record.get("Date", "")) or ts.strftime("%Y-%m-%d")
        if record_date != today_str:
            continue
        rows.append((ts, action, clean_text(record.get("Project", ""))))

    rows.sort(key=lambda r: r[0])
    times = {}
    last_project = None
    for ts, action, project in rows:
        if action in ("Check In", "Lunch Start"):
            times.setdefault(action, ts)
        elif action in ("Lunch End", "Check Out"):
            times[action] = ts
        if project:
            last_project = project

    return {action: format_clock(ts) for action, ts in times.items()}, last_project


def suggest_next_action(done):
    if "Check In" not in done:
        return "Check In"
    if "Check Out" in done:
        return None
    if "Lunch Start" not in done:
        return "Lunch Start"
    if "Lunch End" not in done:
        return "Lunch End"
    return "Check Out"


def render_today_steps(done, next_action):
    cells = []
    for action in ATTENDANCE_ACTIONS:
        if action in done:
            css, value = "done", f"✓ {done[action]}"
        elif action == next_action:
            css, value = "next", "Next"
        else:
            css, value = "", "—"
        cells.append(
            f"<div class='ops-step {css}'>"
            f"<div class='ops-step-label'>{esc(action)}</div>"
            f"<div class='ops-step-value'>{esc(value)}</div>"
            "</div>"
        )
    ops_html("<div class='ops-steps'>" + "".join(cells) + "</div>")


def read_location():
    """Ask the phone/browser for GPS.
    Returns (lat, lon, maps_link, status) where status is ok | waiting | denied | unavailable.
    """
    if get_geolocation is None:
        return "", "", "", "unavailable"

    attempt = st.session_state.get("geo_attempt", 0)
    try:
        try:
            location = get_geolocation(component_key=f"geo_{attempt}")
        except TypeError:
            location = get_geolocation()
    except Exception:
        return "", "", "", "denied"

    if location is None:
        # First run: the browser hasn't answered yet. The component reruns the page when it does.
        return "", "", "", "waiting"

    coords = location.get("coords") if isinstance(location, dict) else None
    if isinstance(coords, dict):
        lat = coords.get("latitude", "")
        lon = coords.get("longitude", "")
        if lat not in ("", None) and lon not in ("", None):
            return lat, lon, f"https://www.google.com/maps?q={lat},{lon}", "ok"

    return "", "", "", "denied"


def render_location_status(location_link, status):
    if status == "ok":
        ops_html(
            "<span class='ops-pill ok'>📍 Location detected</span>"
            f"<a class='ops-link' href='{esc(location_link)}' target='_blank'>View on map</a>"
        )
        return

    if status == "unavailable":
        st.error("The location tool (streamlit_js_eval) is not installed, so attendance can't be registered.")
        return

    if status == "waiting":
        ops_html("<span class='ops-pill warn'>📍 Getting your location…</span>")
        st.caption("If your phone asks for permission, tap Allow. This can take a few seconds.")
    else:
        ops_html("<span class='ops-pill bad'>📍 Location not available</span>")
        st.caption("Location is blocked or your phone couldn't get a GPS signal.")

    with st.expander("How to allow location"):
        st.markdown(
            "1. Tap the lock / ⓘ icon next to the website address.\n"
            "2. Open **Site settings** or **Permissions** and set **Location** to **Allow**.\n"
            "3. Check that Location / GPS is turned on in your phone settings.\n"
            "4. Tap **Try again**."
        )
    if st.button("🔄 Try again", key="geo_retry"):
        st.session_state["geo_attempt"] = st.session_state.get("geo_attempt", 0) + 1
        st.rerun()


def style_daily_report(df):
    """Highlight days that need review. Falls back to a plain table if styling isn't available."""
    try:
        colors = {
            REVIEW_STATUS: "background-color: rgba(239, 68, 68, 0.14)",
            EDITED_STATUS: "background-color: rgba(59, 130, 246, 0.12)",
        }

        def highlight(row):
            return [colors.get(str(row.get("Status", "")), "")] * len(row)

        return df.style.apply(highlight, axis=1).format({"Lunch Hours": "{:.2f}", "Payable Hours": "{:.2f}"})
    except Exception:
        return df


# =========================================================
# MANUAL ATTENDANCE CORRECTIONS
# =========================================================

def ensure_note_column(attendance_ws, values):
    """Make sure the Attendance sheet has a 'Note' header. Returns its 0-based column index."""
    header = [h.strip() for h in values[0]] if values else []
    if "Note" in header:
        return header.index("Note")
    col = max((i for i, h in enumerate(header) if h), default=-1) + 1
    col_count = getattr(attendance_ws, "col_count", None)
    if isinstance(col_count, int) and col + 1 > col_count:
        attendance_ws.add_cols(col + 1 - col_count)
    attendance_ws.update_cell(1, col + 1, "Note")
    return col


def find_day_rows(values, worker, date_str):
    """Rows (sheet row number, padded row) of one worker on one work date, read from the live sheet."""
    header = [h.strip() for h in values[0]]
    idx = {h: i for i, h in enumerate(header) if h}
    found = []
    for row_number, row in enumerate(values[1:], start=2):
        row = row + [""] * (len(header) - len(row))
        if row[idx["Name"]].strip() != worker:
            continue
        row_date = row[idx["Date"]].strip()
        if row_date != date_str:
            ts = parse_sheet_timestamp(row[idx["Timestamp"]])
            parsed_date = pd.to_datetime(row_date, errors="coerce") if row_date else pd.NaT
            same_day = (not pd.isna(parsed_date) and parsed_date.strftime("%Y-%m-%d") == date_str) or (
                pd.isna(parsed_date) and ts is not None and ts.strftime("%Y-%m-%d") == date_str
            )
            if not same_day:
                continue
        found.append((row_number, row))
    return found, idx


def apply_attendance_correction(attendance_ws, worker, work_date, project, targets, reason, void_actions):
    """Void the current rows of `void_actions` and append manual rows for `targets` ({action: datetime}).
    Nothing is deleted: replaced rows keep their data with ' (void)' added to the action.
    """
    values = attendance_ws.get_all_values()
    if not values:
        raise ValueError("The Attendance sheet is empty.")

    note_idx = ensure_note_column(attendance_ws, values)
    date_str = work_date.strftime("%Y-%m-%d")
    rows, idx = find_day_rows(values, worker, date_str)
    stamp = melbourne_now().strftime("%d/%m/%Y %H:%M")

    for row_number, row in rows:
        action = row[idx["Action"]].strip()
        if action in void_actions:
            attendance_ws.update_cell(row_number, idx["Action"] + 1, action + VOID_SUFFIX)
            attendance_ws.update_cell(row_number, note_idx + 1, f"Replaced by manual correction {stamp}: {reason}")

    entries = [(worker, work_date, project, action, targets[action]) for action in ATTENDANCE_ACTIONS if targets.get(action)]
    if entries:
        attendance_ws.append_rows(
            [build_manual_row(values[0], idx, note_idx, entry, f"Manual correction {stamp}: {reason}") for entry in entries],
            value_input_option="RAW"
        )


def build_manual_row(header, idx, note_idx, entry, note):
    worker, work_date, project, action, when = entry
    row = [""] * max(len(header), note_idx + 1)
    row[idx.get("Timestamp", 0)] = when.strftime("%Y-%m-%d %H:%M:%S")
    row[idx.get("Date", 1)] = work_date.strftime("%Y-%m-%d")   # stays on the work date, even for a night-shift check out
    row[idx.get("Name", 2)] = worker
    row[idx.get("Project", 3)] = project
    row[idx.get("Action", 4)] = action
    if "Location Link" in idx:
        row[idx["Location Link"]] = MANUAL_SOURCE
    row[note_idx] = note
    return row


def add_manual_rows(attendance_ws, entries, reason):
    """Append several manual records at once (one read + one write).
    entries = [(worker, work_date, project, action, datetime), ...]
    """
    values = attendance_ws.get_all_values()
    if not values:
        raise ValueError("The Attendance sheet is empty.")
    note_idx = ensure_note_column(attendance_ws, values)
    header = [h.strip() for h in values[0]]
    idx = {h: i for i, h in enumerate(header) if h}
    stamp = melbourne_now().strftime("%d/%m/%Y %H:%M")
    rows = [build_manual_row(values[0], idx, note_idx, entry, f"Manual correction {stamp}: {reason}") for entry in entries]
    attendance_ws.append_rows(rows, value_input_option="RAW")


def parse_report_date(date_text):
    """'Tue 17/11/2026' -> date"""
    return datetime.strptime(str(date_text).split(" ", 1)[-1], "%d/%m/%Y").date()


def render_quick_lunch(attendance_ws, attendance, daily_report):
    """Add the same lunch break to many days at once (for workers who forgot to mark it)."""
    if daily_report.empty:
        return
    no_lunch = daily_report[daily_report["Note"].astype(str).str.contains("No lunch recorded")]

    with st.container(border=True):
        st.markdown("**🍽️ Add forgotten lunch breaks**")
        if no_lunch.empty:
            st.caption("Every day in this period has a lunch recorded. ✅")
            return

        st.caption(
            f"{len(no_lunch)} day(s) have no lunch recorded. Pick the days, set the lunch time and save. "
            "Days of 5 hours or more are pre-selected."
        )

        options = {}
        defaults = []
        for _, r in no_lunch.iterrows():
            label = f"{r['Worker']} · {r['Date']} · {r['Check In']} to {r['Check Out']} ({float(r['Payable Hours']):.2f} h)"
            options[label] = (r["Worker"], parse_report_date(r["Date"]))
            if float(r["Payable Hours"]) >= 5:
                defaults.append(label)

        with st.form("quick_lunch_form"):
            picked = st.multiselect("Days", list(options), default=defaults)
            c1, c2 = st.columns(2)
            with c1:
                lunch_start = st.time_input("Lunch start", value=datetime.strptime("12:00", "%H:%M").time(), step=300)
            with c2:
                duration = st.selectbox("Duration", [30, 45, 60], format_func=lambda m: f"{m} min")
            reason = st.text_input("Reason (saved in the sheet)", value="Lunch not marked by worker")
            submitted = st.form_submit_button("🍽️ Add lunch to selected days", type="primary", use_container_width=True)

    if not submitted:
        return
    if not picked:
        st.error("Pick at least one day.")
        return
    if lunch_start is None:
        st.error("Set the lunch start time.")
        return

    reason = " ".join(reason.split()) or "Lunch not marked by worker"
    entries = []
    skipped = []
    for label in picked:
        worker, work_date = options[label]
        day = attendance[(attendance["Name"].astype(str) == worker) & (attendance["Date"].dt.date == work_date)]
        check_in = day[day["Action"] == "Check In"]["Timestamp"]
        check_out = day[day["Action"] == "Check Out"]["Timestamp"]
        if check_in.empty or check_out.empty:
            skipped.append(label.split(" · ")[0] + " " + label.split(" · ")[1])
            continue

        start = datetime.combine(work_date, lunch_start)
        if start <= check_in.iloc[0]:
            start += timedelta(days=1)        # night shift: lunch after midnight
        end = start + timedelta(minutes=int(duration))
        if not (check_in.iloc[0] < start and end < check_out.iloc[-1]):
            skipped.append(label.split(" · ")[0] + " " + label.split(" · ")[1])
            continue

        project = str(day["Project"].iloc[-1]).strip()
        entries.append((worker, work_date, project, "Lunch Start", start))
        entries.append((worker, work_date, project, "Lunch End", end))

    if not entries:
        st.error("The lunch time doesn't fit between Check In and Check Out on the selected days.")
        return

    try:
        add_manual_rows(attendance_ws, entries, reason)
    except Exception as e:
        st.error("Could not save the lunch breaks.")
        st.exception(e)
        return

    clear_sheet_cache()
    message = f"Lunch added to {len(entries) // 2} day(s)."
    if skipped:
        message += " Skipped (lunch outside working hours): " + ", ".join(skipped) + "."
    st.session_state["payroll_flash"] = message
    st.rerun()


def render_fix_attendance(attendance_ws, attendance_all, daily_report, worker_options, default_date):
    st.markdown("**✏️ Fix one day**")
    st.caption(
        "Add or change any time without opening Google Sheets. "
        "The original records stay in the sheet, marked as replaced, with your reason."
    )

    other_label = "➕ Another worker or day"
    review_days = {}
    if not daily_report.empty:
        for _, r in daily_report[daily_report["Status"] == REVIEW_STATUS].iterrows():
            review_days[f"{r['Worker']} · {r['Date']} · {r['Note']}"] = (r["Worker"], parse_report_date(r["Date"]))
        no_lunch = daily_report[
            (daily_report["Status"] != REVIEW_STATUS)
            & daily_report["Note"].astype(str).str.contains("No lunch recorded")
        ]
        for _, r in no_lunch.iterrows():
            review_days[f"{r['Worker']} · {r['Date']} · No lunch recorded"] = (r["Worker"], parse_report_date(r["Date"]))

    choice = st.selectbox("Day to fix", list(review_days) + [other_label])

    if choice == other_label:
        c1, c2 = st.columns(2)
        with c1:
            fix_worker = st.selectbox("Worker", worker_options)
        with c2:
            fix_date = st.date_input("Date", value=default_date, format="DD/MM/YYYY")
    else:
        fix_worker, fix_date = review_days[choice]

    if not fix_worker or not fix_date:
        return

    day = attendance_all[
        (attendance_all["Name"].astype(str).str.strip() == fix_worker)
        & (attendance_all["Date"].dt.date == fix_date)
    ].sort_values("Timestamp")

    st.markdown(f"**Current records · {esc(fix_worker)} · {fix_date.strftime('%a %d/%m/%Y')}**")
    if day.empty:
        st.info("No records for this day yet. Enter the times below to add the day.")
    else:
        links = day["Location Link"].astype(str) if "Location Link" in day.columns else pd.Series("", index=day.index)
        history = pd.DataFrame({
            "Time": day["Timestamp"].dt.strftime("%a %I:%M %p"),
            "Action": day["Action"].astype(str),
            "Project": day["Project"].astype(str),
            "Source": ["Manual" if link == MANUAL_SOURCE else ("GPS" if link.strip() else "—") for link in links],
            "Note": day["Note"].astype(str) if "Note" in day.columns else "",
        })
        st.dataframe(history, use_container_width=True, hide_index=True)

    active = day[day["Action"].isin(ATTENDANCE_ACTIONS)]
    current = {}
    for action in ATTENDANCE_ACTIONS:
        stamps = active[active["Action"] == action]["Timestamp"]
        if not stamps.empty:
            current[action] = (stamps.iloc[0] if action in ("Check In", "Lunch Start") else stamps.iloc[-1]).to_pydatetime()

    current_project = str(active["Project"].iloc[-1]).strip() if not active.empty else ""
    project_choices = get_project_options(None)
    if current_project and current_project not in project_choices:
        project_choices = [current_project] + project_choices
    checkout_next_day = "Check Out" in current and current["Check Out"].date() > fix_date

    form_id = f"fix_{fix_worker}_{fix_date.isoformat()}"
    with st.form(form_id):
        st.markdown("**Correct times**")
        time_cols = st.columns(4)
        new_times = {}
        for col, action in zip(time_cols, ATTENDANCE_ACTIONS):
            with col:
                value = current[action].time().replace(second=0, microsecond=0) if action in current else None
                new_times[action] = st.time_input(action, value=value, step=300, key=f"{form_id}_{action}")

        night_shift = st.checkbox(
            "Check Out is after midnight (night shift)",
            value=checkout_next_day,
            key=f"{form_id}_night"
        )

        p1, p2 = st.columns(2)
        with p1:
            if project_choices:
                default_index = project_choices.index(current_project) if current_project in project_choices else 0
                fix_project = st.selectbox("Project", project_choices, index=default_index, key=f"{form_id}_project")
            else:
                fix_project = st.text_input("Project", value=current_project, key=f"{form_id}_project_text").strip()
        with p2:
            remove = st.multiselect(
                "Remove times",
                [a for a in ATTENDANCE_ACTIONS if a in current],
                key=f"{form_id}_remove"
            )

        reason = st.text_input(
            "Reason (saved in the sheet)",
            placeholder="Forgot to check out, confirmed with supervisor",
            key=f"{form_id}_reason"
        )
        delete_day = st.checkbox("🗑️ Delete this whole day (all records are marked as replaced)", key=f"{form_id}_delete")
        save = st.form_submit_button("💾 Save correction", type="primary", use_container_width=True)

    if not save:
        return

    reason = " ".join(reason.split())
    if not reason:
        st.error("Add a short reason so Aaron can see why it was changed.")
        return
    if not fix_project and not delete_day:
        st.error("Choose the project.")
        return

    if delete_day:
        if not current:
            st.info("There's nothing to delete for this day.")
            return
        void_actions = set(current)
        targets = {}
    else:
        targets_all = {}
        for action in ATTENDANCE_ACTIONS:
            t = None if action in remove else new_times[action]
            if t is None:
                continue
            when = datetime.combine(fix_date, t)
            if action == "Check Out" and night_shift:
                when += timedelta(days=1)
            targets_all[action] = when

        ordered = [(a, targets_all[a]) for a in ATTENDANCE_ACTIONS if a in targets_all]
        for (first, t1), (second, t2) in zip(ordered, ordered[1:]):
            if t2 <= t1:
                st.error(f"{second} must be after {first}. Tick the night-shift box if the check out is after midnight.")
                return

        if ("Lunch Start" in targets_all) != ("Lunch End" in targets_all):
            st.error("Enter both Lunch Start and Lunch End, or remove both.")
            return

        if fix_date < melbourne_now().date() and ("Check In" not in targets_all or "Check Out" not in targets_all):
            st.error("Check In and Check Out are both needed. To remove the day, tick 'Delete this whole day'.")
            return

        def minute(value):
            return value.strftime("%Y-%m-%d %H:%M") if value is not None else None

        changed = {a for a in ATTENDANCE_ACTIONS if minute(current.get(a)) != minute(targets_all.get(a))}
        if current and fix_project != current_project:
            changed |= set(current) | set(targets_all)

        if not changed:
            st.info("Nothing changed.")
            return

        void_actions = {a for a in changed if a in current}
        targets = {a: targets_all[a] for a in changed if a in targets_all}

    try:
        apply_attendance_correction(attendance_ws, fix_worker, fix_date, fix_project, targets, reason, void_actions)
    except Exception as e:
        st.error("Could not save the correction.")
        st.exception(e)
        return

    clear_sheet_cache()
    st.session_state["payroll_flash"] = (
        f"Day deleted for {fix_worker} on {fix_date.strftime('%d/%m/%Y')}."
        if delete_day
        else f"Correction saved for {fix_worker} on {fix_date.strftime('%d/%m/%Y')}."
    )
    st.rerun()


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
        escape=True
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
        parts.append(f"<h3>{esc(worker)}</h3>")
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

    first_name = esc(str(worker_name).split()[0]) if str(worker_name).strip() else "there"

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
    inject_ops_css()
    ops_hero("👷 Attendance", f"Aaron's Demolition · {melbourne_now().strftime('%A %d %B %Y')}")

    try:
        spreadsheet = connect_spreadsheet()
        workers_ws = get_or_create_worksheet(spreadsheet, WORKERS_SHEET_NAME, WORKERS_HEADERS)
        attendance_ws = get_or_create_worksheet(spreadsheet, ATTENDANCE_SHEET_NAME, ATTENDANCE_HEADERS)
        projects_ws = get_or_create_worksheet(spreadsheet, PROJECTS_SHEET_NAME, PROJECTS_HEADERS)
    except Exception as e:
        st.error("Could not connect to Google Sheets.")
        st.exception(e)
        return

    # Confirmation from the last registration (shown after the page refreshes).
    flash = st.session_state.pop("attendance_flash", None)
    if flash:
        note = f"<div style='margin-top:4px;opacity:0.85'>{esc(flash['note'])}</div>" if flash.get("note") else ""
        ops_html(
            "<div class='ops-success'>"
            f"<div class='ops-success-title'>✅ {esc(flash['action'])} registered</div>"
            f"<div>{esc(flash['name'])} · {esc(flash['time'])} · {esc(flash['project'])}</div>"
            f"{note}</div>"
        )

    lat, lon, location_link, loc_status = read_location()
    render_location_status(location_link, loc_status)

    workers_data = get_sheet_records_cached(WORKERS_SHEET_NAME, tuple(WORKERS_HEADERS))
    all_worker_names = []
    active_names = set()
    for worker in workers_data:
        worker_name = clean_text(worker.get("Name", ""))
        if not worker_name:
            continue
        all_worker_names.append(worker_name)
        if clean_text(worker.get("Active", "Yes")).lower() not in INACTIVE_VALUES:
            active_names.add(worker_name)
    worker_names = sorted(active_names)

    # After a first-time registration, jump straight to the worker's own screen.
    pending_worker = st.session_state.pop("_pending_worker", None)
    if pending_worker:
        st.session_state["register_type"] = REGISTER_OPTIONS[0]
        if pending_worker in worker_names:
            st.session_state["attendance_worker"] = pending_worker

    if st.session_state.get("attendance_worker") not in [None] + worker_names:
        del st.session_state["attendance_worker"]

    register_type = st.radio(
        "Registration",
        REGISTER_OPTIONS,
        horizontal=True,
        key="register_type",
        label_visibility="collapsed"
    )

    location = (lat, lon, location_link, loc_status)

    if register_type == REGISTER_OPTIONS[0]:
        attendance_existing_worker(worker_names, attendance_ws, projects_ws, location)
    else:
        attendance_first_time(all_worker_names, workers_ws, attendance_ws, projects_ws, location)


def location_not_ready_message():
    st.error("📍 Your location isn't ready yet. Allow location access, wait a few seconds and tap the button again.")


def attendance_existing_worker(worker_names, attendance_ws, projects_ws, location):
    lat, lon, location_link, loc_status = location

    if not worker_names:
        st.warning("No registered workers yet. Choose **First time here** above.")
        return

    with st.container(border=True):
        name = st.selectbox(
            "👤 Your name",
            worker_names,
            index=None,
            placeholder="Select your name",
            key="attendance_worker"
        )

        if not name:
            st.caption("Select your name to see today's attendance.")
            return

        records = get_sheet_records_cached(ATTENDANCE_SHEET_NAME, tuple(ATTENDANCE_HEADERS))
        today_str = melbourne_now().strftime("%Y-%m-%d")
        done, last_project = todays_attendance(records, name, today_str)
        next_action = suggest_next_action(done)

        project = project_picker(projects_ws, default_project=last_project, key_prefix="att")

        st.markdown("**Today**")
        render_today_steps(done, next_action)
        if next_action is None:
            ops_html("<span class='ops-pill ok'>Day complete ✅</span>")

        # Pre-select the next logical step; the worker can still choose another one.
        default_marker = (name, today_str, tuple(sorted(done.items())))
        if st.session_state.get("_attendance_default_for") != default_marker:
            st.session_state["_attendance_default_for"] = default_marker
            st.session_state["attendance_action"] = next_action or "Check Out"

        action = st.radio(
            "What are you registering?",
            ATTENDANCE_ACTIONS,
            horizontal=True,
            key="attendance_action"
        )

        if action in done:
            st.warning(f"You already registered **{action}** today at {done[action]}.")

        if st.button(f"✅ Register {action}", type="primary", use_container_width=True, key="register_attendance"):
            if loc_status != "ok":
                location_not_ready_message()
                return
            if not project:
                st.error("Please choose or type the project.")
                return

            now = melbourne_now()
            try:
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
            except Exception as e:
                st.error("Could not save your attendance. Please try again.")
                st.exception(e)
                return

            clear_sheet_cache()
            st.session_state["attendance_flash"] = {
                "action": action,
                "name": name,
                "time": format_clock(now),
                "project": project,
            }
            st.rerun()


def attendance_first_time(all_worker_names, workers_ws, attendance_ws, projects_ws, location):
    lat, lon, location_link, loc_status = location

    with st.container(border=True):
        st.markdown("**Your details**")
        st.caption("Your bank details and email are saved only once. Next time choose **I'm already registered**.")

        with st.form("first_time_registration"):
            c1, c2 = st.columns(2)
            with c1:
                first_name = st.text_input("First name")
            with c2:
                last_name = st.text_input("Last name")

            email = st.text_input("Email")

            c3, c4 = st.columns(2)
            with c3:
                bsb = st.text_input("BSB", placeholder="6 digits, e.g. 063-000")
            with c4:
                account = st.text_input("Account number")

            project = project_picker(projects_ws, key_prefix="reg")
            st.caption("Registering also records your **Check In** for today.")

            submitted = st.form_submit_button("✅ Register and Check In", type="primary", use_container_width=True)

    if not submitted:
        return

    full_name = " ".join(f"{first_name} {last_name}".split())
    email_clean = clean_text(email)
    bsb_digits = normalize_account(bsb)        # digits only, no padding
    account_digits = normalize_account(account)

    problems = []
    if not first_name.strip() or not last_name.strip():
        problems.append("first and last name")
    if "@" not in email_clean or "." not in email_clean.split("@")[-1]:
        problems.append("a valid email")
    if len(bsb_digits) != 6:
        problems.append("a BSB with exactly 6 digits")
    if not 5 <= len(account_digits) <= 10:
        problems.append("an account number (5 to 10 digits)")
    if not project:
        problems.append("the project")

    if problems:
        st.error("Please check: " + ", ".join(problems) + ".")
        return

    if loc_status != "ok":
        location_not_ready_message()
        return

    existing = {n.lower(): n for n in all_worker_names}
    already_registered = full_name.lower() in existing
    if already_registered:
        full_name = existing[full_name.lower()]

    now = melbourne_now()
    try:
        if not already_registered:
            workers_ws.append_row([
                full_name,
                35,
                bsb_digits,
                account_digits,
                0,
                "Yes",
                email_clean
            ], value_input_option="RAW")

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
    except Exception as e:
        st.error("Could not save your registration. Please try again.")
        st.exception(e)
        return

    clear_sheet_cache()
    st.session_state["attendance_flash"] = {
        "action": "Check In",
        "name": full_name,
        "time": format_clock(now),
        "project": project,
        "note": (
            "You were already registered, so your bank details were not changed."
            if already_registered
            else "Registration saved. Next time choose “I'm already registered”."
        ),
    }
    st.session_state["_pending_worker"] = full_name
    st.rerun()


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


def compute_payroll(filtered, workers):
    """Turn attendance rows into a per-worker payroll summary and a daily breakdown."""
    results = []
    daily_rows = []

    def clock(ts):
        return ts.strftime("%I:%M %p")

    for name in sorted(filtered["Name"].dropna().astype(str).unique()):
        worker_records = filtered[filtered["Name"].astype(str) == name].sort_values("Timestamp")
        total_normal_hours = 0.0
        total_sunday_hours = 0.0
        total_lunch_hours = 0.0

        for work_date in sorted(worker_records["Date"].dt.date.unique()):
            day_records = worker_records[worker_records["Date"].dt.date == work_date].sort_values("Timestamp")
            projects_text = ", ".join(sorted(day_records["Project"].dropna().astype(str).unique()))
            date_text = work_date.strftime("%a %d/%m/%Y")
            was_edited = (
                "Location Link" in day_records.columns
                and (day_records["Location Link"].astype(str) == MANUAL_SOURCE).any()
            )
            edited_note = " - manual correction" if was_edited else ""

            check_in = day_records[day_records["Action"] == "Check In"]["Timestamp"]
            lunch_start = day_records[day_records["Action"] == "Lunch Start"]["Timestamp"]
            lunch_end = day_records[day_records["Action"] == "Lunch End"]["Timestamp"]
            check_out = day_records[day_records["Action"] == "Check Out"]["Timestamp"]

            if check_in.empty or check_out.empty:
                missing = []
                if check_in.empty:
                    missing.append("check in")
                if check_out.empty:
                    missing.append("check out")

                daily_rows.append({
                    "Worker": name,
                    "Date": date_text,
                    "Project": projects_text,
                    "Check In": clock(check_in.iloc[0]) if not check_in.empty else "Missing",
                    "Lunch Start": clock(lunch_start.iloc[0]) if not lunch_start.empty else "",
                    "Lunch End": clock(lunch_end.iloc[-1]) if not lunch_end.empty else "",
                    "Check Out": clock(check_out.iloc[-1]) if not check_out.empty else "Missing",
                    "Lunch Hours": 0.0,
                    "Payable Hours": 0.0,
                    "Status": REVIEW_STATUS,
                    "Note": "Missing " + " and ".join(missing) + " - not paid" + edited_note
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
                lunch_start_text = clock(lunch_s)
                lunch_end_text = clock(lunch_e)
                lunch_note = "Lunch deducted"
            elif not lunch_start.empty or not lunch_end.empty:
                lunch_note = "Incomplete lunch record"
                if not lunch_start.empty:
                    lunch_start_text = clock(lunch_start.iloc[0])
                if not lunch_end.empty:
                    lunch_end_text = clock(lunch_end.iloc[-1])

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
                "Date": date_text,
                "Project": projects_text,
                "Check In": clock(start),
                "Lunch Start": lunch_start_text,
                "Lunch End": lunch_end_text,
                "Check Out": clock(end),
                "Lunch Hours": round(lunch_hours, 2),
                "Payable Hours": round(payable_hours, 2),
                "Status": (
                    REVIEW_STATUS if lunch_note == "Incomplete lunch record"
                    else EDITED_STATUS if was_edited
                    else OK_STATUS
                ),
                "Note": f"{pay_note} - {lunch_note}{edited_note}"
            })

        worker_row = workers[workers["Name"].astype(str).str.strip() == name] if "Name" in workers.columns else workers.iloc[0:0]
        rate = 35.0
        sunday_rate = 0.0
        bsb = ""
        account = ""
        email = ""

        if not worker_row.empty:
            rate = safe_float(worker_row.iloc[0].get("Rate", 35), 35.0)
            sunday_rate = safe_float(worker_row.iloc[0].get("Sunday Rate", 0), 0.0)
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

    return pd.DataFrame(results), pd.DataFrame(daily_rows)


def payroll_page():
    inject_ops_css()
    ops_hero(
        "💵 Weekly Payroll",
        "Weeks run Monday to Sunday. Review hours, rates and adjustments, then export or email."
    )

    try:
        spreadsheet = connect_spreadsheet()
        attendance_ws = get_or_create_worksheet(spreadsheet, ATTENDANCE_SHEET_NAME, ATTENDANCE_HEADERS)
    except Exception as e:
        st.error("Could not connect to Google Sheets.")
        st.exception(e)
        return

    flash = st.session_state.pop("payroll_flash", None)
    if flash:
        st.success(f"✅ {flash}")

    workers = pd.DataFrame(get_sheet_records_cached(WORKERS_SHEET_NAME, tuple(WORKERS_HEADERS)))
    attendance = pd.DataFrame(get_sheet_records_cached(ATTENDANCE_SHEET_NAME, tuple(ATTENDANCE_HEADERS)))

    if workers.empty:
        st.warning("No workers found. Workers must register first from the QR page.")
        return

    if attendance.empty:
        st.info("No attendance records yet.")
        return

    required_cols = ["Timestamp", "Date", "Name", "Project", "Action"]
    missing_cols = [c for c in required_cols if c not in attendance.columns]
    if missing_cols:
        st.error(f"Missing columns in Attendance sheet: {missing_cols}")
        return

    attendance["Timestamp"] = pd.to_datetime(attendance["Timestamp"], errors="coerce")
    attendance["Date"] = pd.to_datetime(attendance["Date"], errors="coerce")
    attendance = attendance.dropna(subset=["Timestamp", "Date"])
    attendance["Name"] = attendance["Name"].astype(str).str.strip()

    # Keep every row for the correction history, but leave replaced (void) rows out of payroll.
    attendance_all = attendance.copy()
    attendance = attendance[attendance["Action"].astype(str).isin(ATTENDANCE_ACTIONS)].copy()

    if attendance.empty:
        st.warning("No valid attendance records found.")
        return

    week_options = get_week_options(attendance["Date"])
    week_labels = [w[0] for w in week_options]

    # ---------------- Filters ----------------
    with st.container(border=True):
        st.markdown("**Payroll period**")
        f1, f2, f3, f4 = st.columns([2, 2, 2, 1])

        with f1:
            start_week_label = st.selectbox("Start week", week_labels, index=0)
        with f2:
            end_week_label = st.selectbox("End week", week_labels, index=0)

        project_options = ["All Projects"] + sorted(attendance["Project"].dropna().astype(str).unique().tolist())
        with f3:
            selected_project = st.selectbox("Project", project_options)
        with f4:
            ops_spacer()
            if st.button("🔄 Refresh", use_container_width=True, help="Reload the latest data from Google Sheets"):
                clear_sheet_cache()
                st.rerun()

        start_week = next(w for w in week_options if w[0] == start_week_label)
        end_week = next(w for w in week_options if w[0] == end_week_label)
        start_date = min(start_week[1], end_week[1])
        end_date = max(start_week[2], end_week[2])
        selected_label = f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}"

        filtered = attendance[
            (attendance["Date"].dt.date >= start_date) &
            (attendance["Date"].dt.date <= end_date)
        ].copy()

        if selected_project != "All Projects":
            filtered = filtered[filtered["Project"].astype(str) == selected_project]

        if filtered.empty:
            st.info(f"No attendance records for {selected_label}" + ("" if selected_project == "All Projects" else f" on {selected_project}") + ".")
            return

        names = sorted(filtered["Name"].dropna().astype(str).unique().tolist())
        w1, w2 = st.columns([1, 3])
        with w1:
            report_type = st.radio("Workers", ["All workers", "Selected only"])
        with w2:
            if report_type == "Selected only":
                selected_workers = st.multiselect("Select workers to include", names, default=names)
            else:
                selected_workers = names
                ops_spacer("2rem")
                st.caption(f"{len(names)} worker(s) with attendance in {selected_label}.")

    if not selected_workers:
        st.warning("No workers selected for this report.")
        return

    filtered = filtered[filtered["Name"].astype(str).isin(selected_workers)]
    report, daily_report = compute_payroll(filtered, workers)
    report = report.reset_index(drop=True)
    report["Adjustment AUD"] = 0.0

    kpi_area = st.container()
    alert_area = st.container()

    tab_summary, tab_daily, tab_fix, tab_send = st.tabs(
        ["💵 Payroll summary", "📅 Daily attendance", "✏️ Fix attendance", "📤 Export & email"]
    )

    editor_signature = f"{start_date}|{end_date}|{selected_project}|{'|'.join(report['Name'])}"
    editor_key = "payroll_editor_" + hashlib.md5(editor_signature.encode("utf-8")).hexdigest()[:10]

    # ---------------- Summary tab ----------------
    with tab_summary:
        st.caption(
            "Edit **Rate**, **Sunday Rate** and **Adjustment AUD** (negative for deductions). "
            "The final table below updates automatically."
        )

        edited = st.data_editor(
            report[["Name", "Normal Hours", "Sunday Hours", "Lunch Deducted", "Rate", "Sunday Rate", "Adjustment AUD"]],
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key=editor_key,
            column_config={
                "Name": st.column_config.TextColumn("Name", disabled=True),
                "Normal Hours": st.column_config.NumberColumn("Normal hrs", disabled=True, format="%.2f"),
                "Sunday Hours": st.column_config.NumberColumn("Sunday hrs", disabled=True, format="%.2f"),
                "Lunch Deducted": st.column_config.NumberColumn("Lunch hrs", disabled=True, format="%.2f"),
                "Rate": st.column_config.NumberColumn("Rate ✏️", min_value=0.0, step=1.0, format="$%.2f"),
                "Sunday Rate": st.column_config.NumberColumn("Sunday rate ✏️", min_value=0.0, step=1.0, format="$%.2f"),
                "Adjustment AUD": st.column_config.NumberColumn("Adjustment ✏️", step=10.0, format="$%.2f"),
            }
        )

        editable_report = report.copy()
        for col in ["Rate", "Sunday Rate", "Adjustment AUD"]:
            editable_report[col] = pd.to_numeric(edited[col], errors="coerce").fillna(0).to_numpy()
        for col in ["Normal Hours", "Sunday Hours"]:
            editable_report[col] = pd.to_numeric(editable_report[col], errors="coerce").fillna(0)

        editable_report["Total Hours"] = (editable_report["Normal Hours"] + editable_report["Sunday Hours"]).round(2)
        editable_report["Total Pay"] = (
            editable_report["Normal Hours"] * editable_report["Rate"]
            + editable_report["Sunday Hours"] * editable_report["Sunday Rate"]
            + editable_report["Adjustment AUD"]
        ).round(2)

        st.markdown("##### Final payroll")
        st.dataframe(
            editable_report[[
                "Name", "BSB", "Account", "Normal Hours", "Sunday Hours", "Total Hours",
                "Rate", "Sunday Rate", "Adjustment AUD", "Total Pay"
            ]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "BSB": st.column_config.TextColumn("BSB"),
                "Account": st.column_config.TextColumn("Account"),
                "Normal Hours": st.column_config.NumberColumn("Normal hrs", format="%.2f"),
                "Sunday Hours": st.column_config.NumberColumn("Sunday hrs", format="%.2f"),
                "Total Hours": st.column_config.NumberColumn("Total hrs", format="%.2f"),
                "Rate": st.column_config.NumberColumn("Rate", format="$%.2f"),
                "Sunday Rate": st.column_config.NumberColumn("Sunday rate", format="$%.2f"),
                "Adjustment AUD": st.column_config.NumberColumn("Adjustment", format="$%.2f"),
                "Total Pay": st.column_config.NumberColumn("Total pay", format="$%.2f"),
            }
        )

    total_workers = len(editable_report)
    total_hours = float(editable_report["Total Hours"].sum()) if not editable_report.empty else 0.0
    total_pay = float(editable_report["Total Pay"].sum()) if not editable_report.empty else 0.0

    review_mask = daily_report["Status"] == REVIEW_STATUS if not daily_report.empty else pd.Series(dtype=bool)
    missing_days = daily_report[daily_report["Note"].astype(str).str.startswith("Missing")] if not daily_report.empty else daily_report
    lunch_issue_days = int(daily_report["Note"].astype(str).str.contains("Incomplete lunch").sum()) if not daily_report.empty else 0
    zero_sunday = editable_report[
        (editable_report["Sunday Hours"] > 0) & (editable_report["Sunday Rate"] <= 0)
    ]["Name"].tolist()

    # ---------------- KPIs + alerts (shown above the tabs) ----------------
    with kpi_area:
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Workers", total_workers)
        m2.metric("Payable hours", f"{total_hours:,.2f}")
        m3.metric("Total payroll", f"AUD {total_pay:,.2f}")
        m4.metric("Days to review", int(review_mask.sum()))

    with alert_area:
        if not missing_days.empty:
            st.warning(
                f"⚠️ {len(missing_days)} day(s) have a missing Check In or Check Out and are paid as 0 hours. "
                "Fix them in the **✏️ Fix attendance** tab."
            )
        if zero_sunday:
            st.warning(
                f"⚠️ Sunday hours with Sunday Rate 0: {', '.join(zero_sunday)}. "
                "Set their Sunday rate in the summary table, otherwise those hours are paid AUD 0."
            )
        if lunch_issue_days:
            st.info(f"ℹ️ {lunch_issue_days} day(s) have an incomplete lunch record, so no lunch was deducted.")
        no_lunch_days = int(
            (daily_report["Note"].astype(str).str.contains("No lunch recorded") & (daily_report["Payable Hours"] >= 5)).sum()
        ) if not daily_report.empty else 0
        if no_lunch_days:
            st.info(
                f"🍽️ {no_lunch_days} day(s) of 5+ hours have no lunch recorded. "
                "Add it in one go in the **✏️ Fix attendance** tab."
            )

    # ---------------- Daily tab ----------------
    with tab_daily:
        d1, d2 = st.columns([2, 1])
        with d1:
            daily_worker = st.selectbox("Worker", ["All workers"] + editable_report["Name"].tolist())
        with d2:
            ops_spacer()
            only_review = st.checkbox("Only days to review")

        view = daily_report.copy()
        if daily_worker != "All workers":
            view = view[view["Worker"] == daily_worker]
        if only_review:
            view = view[view["Status"] == REVIEW_STATUS]

        if view.empty:
            st.success("Nothing to review. ✅" if only_review else "No records.")
        else:
            st.dataframe(style_daily_report(view), use_container_width=True, hide_index=True)
            st.caption("Red = needs review · Blue = corrected manually")

    # ---------------- Fix attendance tab ----------------
    with tab_fix:
        worker_options = sorted(
            set(attendance_all["Name"].dropna().astype(str))
            | {
                clean_text(n) for n, a in zip(workers.get("Name", []), workers.get("Active", ["Yes"] * len(workers)))
                if clean_text(n) and clean_text(a).lower() not in INACTIVE_VALUES
            }
        )
        default_fix_date = min(end_date, melbourne_now().date())
        render_quick_lunch(attendance_ws, attendance, daily_report)
        render_fix_attendance(attendance_ws, attendance_all, daily_report, worker_options, default_fix_date)

    # ---------------- Export & email tab ----------------
    payroll_filename = f"payroll_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
    report_for_aaron = editable_report.drop(columns=["Email"], errors="ignore")
    excel_bytes = make_payroll_excel(report_for_aaron, daily_report, selected_label)

    with tab_send:
        st.download_button(
            "⬇️ Download payroll Excel",
            data=excel_bytes,
            file_name=payroll_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

        a_col, w_col = st.columns(2)

        with a_col:
            with st.container(border=True):
                st.markdown("#### 📧 Send to Aaron")
                default_receiver = str(get_secret_value("EMAIL_RECEIVER", "")).strip()
                email_to = st.text_input("Aaron's email", value=default_receiver)
                email_subject = st.text_input(
                    "Subject",
                    value=f"Weekly Payroll Report - {selected_label}",
                    key=f"aaron_subject_{selected_label}"
                )

                aaron_email_preview = build_payroll_email_html(
                    report_for_aaron,
                    daily_report,
                    selected_label,
                    total_hours,
                    total_pay,
                    custom_message=""
                )

                with st.expander("👁️ Preview email"):
                    components.html(aaron_email_preview, height=520, scrolling=True)

                if st.button("📤 Send payroll to Aaron", type="primary", use_container_width=True, key="send_payroll_aaron"):
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
                            st.success("Payroll report sent to Aaron.")
                        except Exception as e:
                            st.error("Could not send payroll email to Aaron.")
                            st.exception(e)

        with w_col:
            with st.container(border=True):
                st.markdown("#### 👷 Send to a worker")
                worker_email_options = editable_report["Name"].dropna().astype(str).tolist()
                selected_worker_email = st.selectbox("Worker to email", worker_email_options)

                worker_summary = editable_report[editable_report["Name"].astype(str) == selected_worker_email].copy()
                worker_daily = daily_report[daily_report["Worker"].astype(str) == selected_worker_email].copy()

                if not worker_summary.empty:
                    st.caption(
                        f"{worker_summary.iloc[0]['Total Hours']:.2f} hrs · "
                        f"Total pay AUD {worker_summary.iloc[0]['Total Pay']:,.2f}"
                    )

                default_worker_email = ""
                if not worker_summary.empty and "Email" in worker_summary.columns:
                    default_worker_email = str(worker_summary.iloc[0].get("Email", "")).strip()

                worker_to = st.text_input(
                    "Worker's email",
                    value=default_worker_email,
                    key=f"worker_to_{selected_worker_email}"
                )
                worker_subject = st.text_input(
                    "Subject",
                    value=f"Your Attendance Summary - {selected_label}",
                    key=f"worker_subject_{selected_label}_{selected_worker_email}"
                )

                worker_email_preview = build_worker_email_html(
                    selected_worker_email,
                    worker_summary,
                    worker_daily,
                    selected_label,
                    custom_message=""
                )

                with st.expander("👁️ Preview email"):
                    components.html(worker_email_preview, height=520, scrolling=True)

                if st.button("📤 Send to worker", type="primary", use_container_width=True, key="send_worker_email"):
                    if not worker_to.strip():
                        st.warning("Please enter the worker's email address.")
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


def _projects_columns(values):
    header = [h.strip() for h in values[0]] if values else []
    project_idx = header.index("Project") if "Project" in header else 0
    active_idx = header.index("Active") if "Active" in header else 1
    return project_idx, active_idx


def add_project_row(projects_ws, project_name, active):
    """Add a project after checking duplicates against the live sheet. Returns False if it already exists."""
    values = projects_ws.get_all_values()
    project_idx, _ = _projects_columns(values)
    existing = {
        row[project_idx].strip().lower()
        for row in values[1:]
        if len(row) > project_idx and row[project_idx].strip()
    }
    if project_name.lower() in existing:
        return False

    projects_ws.append_row([project_name, "Yes" if active else "No"], value_input_option="RAW")
    return True


def update_project_status(projects_ws, changes):
    """changes = {project_name: True/False}. Updates only the Active cell of those projects."""
    values = projects_ws.get_all_values()
    if not values:
        return
    project_idx, active_idx = _projects_columns(values)

    for row_number, row in enumerate(values[1:], start=2):
        project_name = row[project_idx].strip() if len(row) > project_idx else ""
        if project_name in changes:
            projects_ws.update_cell(row_number, active_idx + 1, "Yes" if changes[project_name] else "No")


def projects_page():
    inject_ops_css()
    ops_hero("🏗️ Projects", "Projects workers can choose when they mark attendance.")

    try:
        spreadsheet = connect_spreadsheet()
        projects_ws = get_or_create_worksheet(spreadsheet, PROJECTS_SHEET_NAME, PROJECTS_HEADERS)
    except Exception as e:
        st.error("Could not connect to Google Sheets.")
        st.exception(e)
        return

    flash = st.session_state.pop("projects_flash", None)
    if flash:
        st.success(flash)

    projects = pd.DataFrame(get_sheet_records_cached(PROJECTS_SHEET_NAME, tuple(PROJECTS_HEADERS)))
    for col in PROJECTS_HEADERS:
        if col not in projects.columns:
            projects[col] = ""
    projects["Project"] = projects["Project"].astype(str).str.strip()
    projects = projects[projects["Project"] != ""].copy()
    projects["Active"] = ~projects["Active"].astype(str).str.strip().str.lower().isin(INACTIVE_VALUES)

    # Activity from attendance: last day used and who is on site today.
    today_str = melbourne_now().strftime("%Y-%m-%d")
    last_activity = {}
    on_site_today = {}
    workers_today = 0

    attendance = pd.DataFrame(get_sheet_records_cached(ATTENDANCE_SHEET_NAME, tuple(ATTENDANCE_HEADERS)))
    if not attendance.empty and {"Project", "Timestamp", "Name", "Action"}.issubset(attendance.columns):
        att = attendance[attendance["Action"].astype(str).isin(ATTENDANCE_ACTIONS)][["Project", "Timestamp", "Name"]].copy()
        att["Project"] = att["Project"].astype(str).str.strip()
        att["Timestamp"] = pd.to_datetime(att["Timestamp"], errors="coerce")
        att = att.dropna(subset=["Timestamp"])
        if not att.empty:
            last_activity = att.groupby("Project")["Timestamp"].max().to_dict()
            today_att = att[att["Timestamp"].dt.strftime("%Y-%m-%d") == today_str]
            on_site_today = today_att.groupby("Project")["Name"].nunique().to_dict()
            workers_today = int(today_att["Name"].nunique())

    active_count = int(projects["Active"].sum())
    inactive_count = int(len(projects) - active_count)

    m1, m2, m3 = st.columns(3)
    m1.metric("Active projects", active_count)
    m2.metric("Inactive projects", inactive_count)
    m3.metric("Workers on site today", workers_today)

    list_col, add_col = st.columns([2, 1])

    # ---------------- Add project ----------------
    with add_col:
        with st.container(border=True):
            st.markdown("#### ➕ Add project")
            with st.form("add_project_form", clear_on_submit=True):
                new_project = st.text_input("Project name", placeholder="e.g. 12 Smith St, Carlton")
                new_active = st.checkbox("Active", value=True)
                add_clicked = st.form_submit_button("➕ Add project", type="primary", use_container_width=True)

            if add_clicked:
                project_name = " ".join(new_project.split())
                if not project_name:
                    st.warning("Please enter a project name.")
                else:
                    try:
                        added = add_project_row(projects_ws, project_name, new_active)
                    except Exception as e:
                        st.error("Could not add the project.")
                        st.exception(e)
                    else:
                        if not added:
                            st.info(f"“{project_name}” already exists.")
                        else:
                            clear_sheet_cache()
                            st.session_state["projects_editor_version"] = st.session_state.get("projects_editor_version", 0) + 1
                            st.session_state["projects_flash"] = f"Project added: {project_name}"
                            st.rerun()

    # ---------------- Project list ----------------
    with list_col:
        with st.container(border=True):
            st.markdown("#### Project list")

            if projects.empty:
                st.info("No projects yet. Add the first one with the form.")
                return

            st.caption("Untick **Active** to hide a project from the workers' list, then click **Save changes**.")

            view = projects[["Project", "Active"]].drop_duplicates(subset=["Project"]).copy()
            view["Workers today"] = view["Project"].map(on_site_today).fillna(0).astype(int)
            view["Last activity"] = view["Project"].map(
                lambda p: last_activity[p].strftime("%d/%m/%Y") if p in last_activity else "—"
            )
            view = view.sort_values(["Active", "Project"], ascending=[False, True]).reset_index(drop=True)

            editor_version = st.session_state.get("projects_editor_version", 0)
            edited = st.data_editor(
                view,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                key=f"projects_editor_{editor_version}",
                column_config={
                    "Project": st.column_config.TextColumn("Project", disabled=True),
                    "Active": st.column_config.CheckboxColumn("Active"),
                    "Workers today": st.column_config.NumberColumn("Workers today", disabled=True),
                    "Last activity": st.column_config.TextColumn("Last activity", disabled=True),
                }
            )

            active_now = edited["Active"].fillna(False).astype(bool).to_numpy()
            changed_mask = view["Active"].to_numpy() != active_now
            changes = dict(zip(view.loc[changed_mask, "Project"], active_now[changed_mask]))
            changes = {project: bool(active) for project, active in changes.items()}

            save_label = f"💾 Save changes ({len(changes)})" if changes else "💾 Save changes"
            if st.button(save_label, type="primary", disabled=not changes, use_container_width=True, key="save_projects"):
                try:
                    update_project_status(projects_ws, changes)
                except Exception as e:
                    st.error("Could not save the changes.")
                    st.exception(e)
                else:
                    clear_sheet_cache()
                    st.session_state["projects_editor_version"] = editor_version + 1
                    st.session_state["projects_flash"] = f"Saved {len(changes)} change(s)."
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

# =========================================================
# ADMIN PASSWORD (everything except the QR attendance screen)
# =========================================================

def admin_gate():
    """Ask for the admin password once per browser session. The QR attendance screen never asks."""
    if st.session_state.get("admin_ok"):
        return

    expected = str(get_secret_value("ADMIN_PASSWORD", "")).strip()

    st.markdown(
        "<style>"
        "[data-testid='stSidebar'],[data-testid='collapsedControl']{display:none;}"
        ".block-container{max-width:480px;padding-top:4rem;}"
        "</style>",
        unsafe_allow_html=True
    )
    inject_ops_css()
    ops_hero("🔒 Aaron's Admin", "Costings, payroll and project settings.")

    if not expected:
        st.error("The admin password is not set up yet.")
        st.caption(
            'In Streamlit Cloud open the app settings → Secrets and add a line: ADMIN_PASSWORD = "your-password". '
            "Save, then reload this page."
        )
        st.stop()

    with st.form("admin_login"):
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in", type="primary", use_container_width=True)

    if submitted:
        if hmac.compare_digest(password.encode("utf-8"), expected.encode("utf-8")):
            st.session_state["admin_ok"] = True
            st.rerun()
        time_mod.sleep(1.5)  # slows down password guessing
        st.error("Wrong password.")

    st.caption("Workers: scan the QR code on site to mark your attendance.")
    st.stop()


admin_gate()

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

if st.sidebar.button("🔒 Log out", use_container_width=True, key="admin_logout"):
    st.session_state.pop("admin_ok", None)
    st.rerun()

if page == "Attendance Register":
    attendance_page()
    st.stop()

if page == "Payroll Report":
    payroll_page()
    st.stop()

if page == "Project Settings":
    projects_page()
    st.stop()

st.title("🧾 Aaron's Pricing Calculator")
st.caption("Demolition, strip-out, flooring and rubbish removal cost calculator.")

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
