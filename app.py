import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from math import ceil
 
# =========================================================
# CONFIG
# =========================================================
 
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQi2XaKTk3vZnIWejQCiCVNqLwJRcDKYThJCKHOH4iPA_JdDQFEUTcKq5BYRDtbAFDYcu6gWnqQgH2/pub?output=csv"
SPREADSHEET_NAME = "Aarons Pricing Database"
COSTING_SHEET_NAME = "Costing_Records"
 
REQUIRED_HEADERS = [
    "quote_id", "date", "client_name", "project_address", "project_type",
    "working_hours", "floor_level", "room_area", "item", "category", "unit",
    "measurement_method", "quantity", "rate", "item_total", "worker_days",
    "equipment_cost", "disposal_cost", "truck_cost", "other_cost",
    "difficulty_percent", "difficulty_allowance", "labour_cost", "supervisor_cost",
    "direct_cost", "minimum_price", "recommended_price", "premium_price",
    "gst_total", "final_total", "notes", "labourers", "estimated_days", "rounded_days"
]
 
st.set_page_config(
    page_title="Aaron's Pricing Calculator",
    page_icon="🧾",
    layout="wide"
)
 
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
        if value is None:
            return 0.0
        clean = str(value)
        clean = clean.replace("AUD", "")
        clean = clean.replace("$", "")
        clean = clean.replace("%", "")
        clean = clean.replace(" ", "")
        clean = clean.strip()
        if "," in clean and "." not in clean:
            clean = clean.replace(",", ".")
        else:
            clean = clean.replace(",", "")
        if clean == "":
            return 0.0
        return float(clean)
    except Exception:
        return 0.0
 
 
def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value)
 
 
def first_non_empty(series, default=""):
    for value in series:
        if str(value).strip() != "":
            return value
    return default
 
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
        ws = spreadsheet.add_worksheet(title=COSTING_SHEET_NAME, rows="5000", cols="80")
        ws.append_row(REQUIRED_HEADERS)
        return ws
 
    ensure_headers(ws)
    return ws
 
 
def ensure_headers(ws):
    values = ws.get_all_values()
    if not values:
        ws.append_row(REQUIRED_HEADERS)
        return
 
    current_headers = values[0]
    updated_headers = list(current_headers)
    changed = False
 
    for header in REQUIRED_HEADERS:
        if header not in updated_headers:
            updated_headers.append(header)
            changed = True
 
    if changed:
        ws.update(range_name="A1", values=[updated_headers])
 
 
def load_saved_costings():
    ws = connect_google_sheet()
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame()
    return pd.DataFrame(records)
 
 
def update_costing_records(quote_id, updated_df):
    ws = connect_google_sheet()
    all_values = ws.get_all_values()
 
    if not all_values:
        return
 
    headers = all_values[0]
    quote_col_index = headers.index("quote_id")
 
    rows_to_keep = [headers]
 
    for row in all_values[1:]:
        row_quote = row[quote_col_index] if len(row) > quote_col_index else ""
        if row_quote != quote_id:
            # Pad row to the header length so Sheets stays aligned.
            padded_row = row + [""] * (len(headers) - len(row))
            rows_to_keep.append(padded_row[:len(headers)])
 
    ws.clear()
    ws.update(range_name="A1", values=rows_to_keep)
 
    final_df = updated_df.copy()
    for header in headers:
        if header not in final_df.columns:
            final_df[header] = ""
    final_df = final_df[headers]
    ws.append_rows(final_df.values.tolist(), value_input_option="RAW")
 
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
 
    for col in ["market_min", "market_recommended", "market_high", "productivity_per_day"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
 
    for col in ["item", "category", "unit", "notes"]:
        df[col] = df[col].astype(str).str.strip()
 
    return df
 
rates_df = load_rates(SHEET_CSV_URL)
 
# =========================================================
# SIDEBAR NAVIGATION
# =========================================================
 
page = st.sidebar.radio(
    "Navigation",
    [
        "New Costing",
        "Saved Costings",
        "Manager View",
        "Edit Costing"
    ]
)
 
st.title("🧾 Aaron's Pricing Calculator")
st.caption("Demolition, strip-out, flooring and rubbish removal cost calculator.")
 
# =========================================================
# SAVED COSTINGS PAGE
# =========================================================
 
if page == "Saved Costings":
    st.title("📂 Saved Costings")
 
    saved_df = load_saved_costings()
 
    if saved_df.empty:
        st.warning("No saved costings found.")
        st.stop()
 
    c1, c2, c3 = st.columns(3)
    with c1:
        client_filter = st.text_input("Search Client")
    with c2:
        project_filter = st.text_input("Search Project Address")
    with c3:
        quote_filter = st.text_input("Search Quote ID")
 
    if client_filter:
        saved_df = saved_df[saved_df["client_name"].astype(str).str.contains(client_filter, case=False, na=False)]
    if project_filter:
        saved_df = saved_df[saved_df["project_address"].astype(str).str.contains(project_filter, case=False, na=False)]
    if quote_filter:
        saved_df = saved_df[saved_df["quote_id"].astype(str).str.contains(quote_filter, case=False, na=False)]
 
    summary_df = saved_df.groupby(
        ["quote_id", "date", "client_name", "project_address"],
        as_index=False
    ).agg({"final_total": "max"})
 
    summary_df["Display"] = (
        summary_df["project_address"].astype(str) + " - " +
        summary_df["client_name"].astype(str) + " (" +
        summary_df["quote_id"].astype(str) + ")"
    )
 
    st.markdown("## Costing Records")
    st.dataframe(summary_df[["quote_id", "date", "client_name", "project_address", "final_total"]], use_container_width=True)
 
    selected_label = st.selectbox("Open Costing", summary_df["Display"].tolist())
    selected_quote = summary_df.loc[summary_df["Display"] == selected_label, "quote_id"].iloc[0]
 
    selected_df = saved_df[saved_df["quote_id"] == selected_quote]
 
    if not selected_df.empty:
        st.markdown("### Full Costing Breakdown")
        st.dataframe(selected_df, use_container_width=True)
 
    st.stop()
 
# =========================================================
# MANAGER VIEW
# =========================================================
 
if page == "Manager View":
    st.title("📊 Manager Costing View")
 
    saved_df = load_saved_costings()
 
    if saved_df.empty:
        st.warning("No saved costings found.")
        st.stop()
 
    summary_df = saved_df.groupby(
        ["quote_id", "date", "client_name", "project_address"],
        as_index=False
    ).agg({"final_total": "max"})
 
    summary_df["Display"] = (
        summary_df["project_address"].astype(str) + " - " +
        summary_df["client_name"].astype(str) + " (" +
        summary_df["quote_id"].astype(str) + ")"
    )
 
    selected_label = st.selectbox("Select Project", summary_df["Display"].tolist())
    selected_quote = summary_df.loc[summary_df["Display"] == selected_label, "quote_id"].iloc[0]
 
    selected_df = saved_df[saved_df["quote_id"] == selected_quote]
 
    if not selected_df.empty:
        first_row = selected_df.iloc[0]
 
        st.markdown("## Project Information")
        c1, c2 = st.columns(2)
        with c1:
            st.metric("Client", clean_text(first_row.get("client_name", "")))
            st.metric("Project Address", clean_text(first_row.get("project_address", "")))
        with c2:
            st.metric("Project Type", clean_text(first_row.get("project_type", "")))
            st.metric("Working Hours", clean_text(first_row.get("working_hours", "")))
 
        st.divider()
 
        st.markdown("## Financial Summary")
        f1, f2, f3 = st.columns(3)
        f1.metric("Direct Cost", clean_text(first_row.get("direct_cost", "")))
        f2.metric("Recommended Price", clean_text(first_row.get("recommended_price", "")))
        f3.metric("Final Total", clean_text(first_row.get("final_total", "")))
 
        f4, f5, f6 = st.columns(3)
        f4.metric("Labour Cost", clean_text(first_row.get("labour_cost", "")))
        f5.metric("Supervisor Cost", clean_text(first_row.get("supervisor_cost", "")))
        f6.metric("Difficulty Allowance", clean_text(first_row.get("difficulty_allowance", "")))
 
        f7, f8, f9 = st.columns(3)
        f7.metric("Equipment Cost", clean_text(first_row.get("equipment_cost", "")))
        f8.metric("Disposal / Tip Fees", clean_text(first_row.get("disposal_cost", "")))
        f9.metric("Truck / Transport", clean_text(first_row.get("truck_cost", "")))
 
        st.divider()
 
        st.markdown("## Program Summary")
        worker_days_total = selected_df["worker_days"].apply(parse_number).sum() if "worker_days" in selected_df.columns else 0
        labourers_value = parse_number(first_row.get("labourers", 0))
        estimated_days_value = parse_number(first_row.get("estimated_days", 0))
        rounded_days_value = parse_number(first_row.get("rounded_days", 0))
 
        if estimated_days_value == 0 and labourers_value > 0:
            estimated_days_value = worker_days_total / labourers_value
        if rounded_days_value == 0 and estimated_days_value > 0:
            rounded_days_value = ceil(estimated_days_value)
 
        p1, p2, p3 = st.columns(3)
        p1.metric("Estimated Worker Days", f"{worker_days_total:.2f}")
        p2.metric("Labourers", f"{int(labourers_value) if labourers_value else 'TBC'}")
        p3.metric("Estimated Duration", f"{estimated_days_value:.2f} days ({int(rounded_days_value)} rounded)")
 
        st.divider()
 
        st.markdown("## Work Breakdown")
        manager_df = selected_df[["room_area", "item", "quantity", "rate", "item_total", "notes"]].copy()
        manager_df.columns = ["Area", "Work Item", "Quantity", "Rate", "Total", "Notes"]
        st.dataframe(manager_df, use_container_width=True)
 
    st.stop()
 
# =========================================================
# EDIT COSTING PAGE
# =========================================================
 
if page == "Edit Costing":
    st.title("✏️ Edit Costing")
 
    saved_df = load_saved_costings()
 
    if saved_df.empty:
        st.warning("No saved costings found.")
        st.stop()
 
    summary_df = saved_df.groupby(
        ["quote_id", "date", "client_name", "project_address"],
        as_index=False
    ).agg({"final_total": "max"})
 
    summary_df["Display"] = (
        summary_df["project_address"].astype(str) + " - " +
        summary_df["client_name"].astype(str) + " (" +
        summary_df["quote_id"].astype(str) + ")"
    )
 
    selected_label = st.selectbox("Select Project to Edit", summary_df["Display"].tolist())
    selected_quote = summary_df.loc[summary_df["Display"] == selected_label, "quote_id"].iloc[0]
 
    selected_df = saved_df[saved_df["quote_id"] == selected_quote].copy()
 
    if selected_df.empty:
        st.warning("No records found for this quote.")
        st.stop()
 
    first_row = selected_df.iloc[0]
 
    st.markdown("## Project Details")
    c1, c2, c3 = st.columns(3)
    with c1:
        edit_client = st.text_input("Client Name", value=clean_text(first_row.get("client_name", "")))
        edit_address = st.text_input("Project Address", value=clean_text(first_row.get("project_address", "")))
    with c2:
        edit_project_type = st.text_input("Project Type", value=clean_text(first_row.get("project_type", "")))
        edit_working_hours = st.text_input("Working Hours", value=clean_text(first_row.get("working_hours", "")))
    with c3:
        edit_floor_level = st.text_input("Floor Level", value=clean_text(first_row.get("floor_level", "")))
 
    st.markdown("## Edit Work Items")
    edit_cols = ["room_area", "item", "category", "unit", "measurement_method", "quantity", "rate", "notes"]
    editable_df = selected_df[edit_cols].copy()
 
    edited_df = st.data_editor(
        editable_df,
        use_container_width=True,
        num_rows="dynamic",
        key="edit_costing_table"
    )
 
    st.markdown("## Recalculate Costs")
    r1, r2, r3, r4 = st.columns(4)
    with r1:
        edit_equipment_cost = st.text_input("Equipment Cost", value=clean_text(first_row.get("equipment_cost", "AUD 0.00")))
    with r2:
        edit_disposal_cost = st.text_input("Disposal / Tip Fees", value=clean_text(first_row.get("disposal_cost", "AUD 0.00")))
    with r3:
        edit_truck_cost = st.text_input("Truck / Transport Cost", value=clean_text(first_row.get("truck_cost", "AUD 0.00")))
    with r4:
        edit_other_cost = st.text_input("Other Cost", value=clean_text(first_row.get("other_cost", "AUD 0.00")))
 
    r5, r6, r7, r8 = st.columns(4)
    with r5:
        edit_labourers = st.number_input("Labourers", min_value=1, value=int(parse_number(first_row.get("labourers", 2)) or 2), step=1)
    with r6:
        edit_labour_cost = st.text_input("Labour Cost", value=clean_text(first_row.get("labour_cost", "AUD 0.00")))
    with r7:
        edit_supervisor_cost = st.text_input("Supervisor Cost", value=clean_text(first_row.get("supervisor_cost", "AUD 0.00")))
    with r8:
        edit_difficulty_percent = st.text_input("Difficulty %", value=clean_text(first_row.get("difficulty_percent", "0.0%")))
 
    r9, r10 = st.columns(2)
    with r9:
        edit_margin = st.number_input("Recommended Margin %", min_value=0.0, max_value=80.0, value=30.0, step=1.0)
    with r10:
        edit_gst = st.number_input("GST %", min_value=0.0, max_value=20.0, value=10.0, step=0.5)
 
    edited_df["quantity_num"] = edited_df["quantity"].apply(parse_number)
    edited_df["rate_num"] = edited_df["rate"].apply(parse_number)
    edited_df["item_total_num"] = edited_df["quantity_num"] * edited_df["rate_num"]
 
    def productivity_for_item(item_name):
        match = rates_df[rates_df["item"] == str(item_name)]
        if match.empty:
            return 0.0
        return float(match.iloc[0]["productivity_per_day"])
 
    edited_df["productivity_num"] = edited_df["item"].apply(productivity_for_item)
    edited_df["worker_days_num"] = edited_df.apply(
        lambda r: r["quantity_num"] / r["productivity_num"] if r["productivity_num"] > 0 else 0,
        axis=1
    )
 
    base_total_edit = edited_df["item_total_num"].sum()
    equipment_num = parse_number(edit_equipment_cost)
    disposal_num = parse_number(edit_disposal_cost)
    truck_num = parse_number(edit_truck_cost)
    other_num = parse_number(edit_other_cost)
    labour_num = parse_number(edit_labour_cost)
    supervisor_num = parse_number(edit_supervisor_cost)
    difficulty_num = parse_number(edit_difficulty_percent)
 
    direct_cost_edit = base_total_edit + equipment_num + disposal_num + truck_num + other_num + labour_num + supervisor_num
    difficulty_allowance_edit = direct_cost_edit * (difficulty_num / 100)
    cost_with_difficulty_edit = direct_cost_edit + difficulty_allowance_edit
 
    recommended_price_edit = cost_with_difficulty_edit / (1 - edit_margin / 100) if 0 < edit_margin < 100 else cost_with_difficulty_edit
    gst_total_edit = recommended_price_edit * (edit_gst / 100)
    final_total_edit = recommended_price_edit + gst_total_edit
 
    worker_days_total_edit = edited_df["worker_days_num"].sum()
    estimated_days_edit = worker_days_total_edit / edit_labourers if edit_labourers > 0 else 0
    rounded_days_edit = ceil(estimated_days_edit) if estimated_days_edit > 0 else 0
 
    st.markdown("## Updated Summary")
    s1, s2, s3 = st.columns(3)
    s1.metric("Direct Cost", money(direct_cost_edit))
    s2.metric("Recommended Price", money(recommended_price_edit))
    s3.metric("Final Total Inc GST", money(final_total_edit))
 
    s4, s5, s6 = st.columns(3)
    s4.metric("Estimated Worker Days", f"{worker_days_total_edit:.2f}")
    s5.metric("Labourers", str(edit_labourers))
    s6.metric("Estimated Duration", f"{estimated_days_edit:.2f} days ({rounded_days_edit} rounded)")
 
    st.markdown("## Save Changes")
 
    if st.button("💾 Update Costing"):
        now = datetime.now().strftime("%d/%m/%Y %H:%M")
        updated_rows = []
 
        for _, row in edited_df.iterrows():
            qty = parse_number(row["quantity"])
            rate = parse_number(row["rate"])
            item_total = qty * rate
            productivity = productivity_for_item(row["item"])
            worker_days = qty / productivity if productivity > 0 else 0
 
            updated_rows.append({
                "quote_id": selected_quote,
                "date": now,
                "client_name": edit_client,
                "project_address": edit_address,
                "project_type": edit_project_type,
                "working_hours": edit_working_hours,
                "floor_level": edit_floor_level,
                "room_area": row["room_area"],
                "item": row["item"],
                "category": row["category"],
                "unit": row["unit"],
                "measurement_method": row["measurement_method"],
                "quantity": round(qty, 2),
                "rate": money(rate),
                "item_total": money(item_total),
                "worker_days": round(worker_days, 2),
                "equipment_cost": money(equipment_num),
                "disposal_cost": money(disposal_num),
                "truck_cost": money(truck_num),
                "other_cost": money(other_num),
                "difficulty_percent": f"{difficulty_num:.1f}%",
                "difficulty_allowance": money(difficulty_allowance_edit),
                "labour_cost": money(labour_num),
                "supervisor_cost": money(supervisor_num),
                "direct_cost": money(direct_cost_edit),
                "minimum_price": "",
                "recommended_price": money(recommended_price_edit),
                "premium_price": "",
                "gst_total": money(gst_total_edit),
                "final_total": money(final_total_edit),
                "notes": row["notes"],
                "labourers": edit_labourers,
                "estimated_days": round(estimated_days_edit, 2),
                "rounded_days": rounded_days_edit
            })
 
        final_df = pd.DataFrame(updated_rows)
        update_costing_records(selected_quote, final_df)
        st.success(f"✅ Quote {selected_quote} updated successfully.")
 
    st.stop()
 
# =========================================================
# NEW COSTING SESSION STATE
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
 
room_options_map = {
    "Residential": ["Kitchen", "Bathroom", "Laundry", "Bedroom", "Living Room", "Dining Room", "Hallway", "Garage", "Backyard", "Other"],
    "Commercial": ["Tenancy Area", "Office", "Reception", "Meeting Room", "Kitchenette", "Bathroom", "Storage", "Shopfront", "Common Area", "Other"],
    "Industrial": ["Warehouse Floor", "Racking Area", "Loading Bay", "Office Area", "Mezzanine", "Yard", "Concrete Area", "Storage Area", "Waste Area", "Other"],
    "Shopping Centre": ["Tenancy Area", "Food Court", "Shopfront", "Kitchen", "Coolroom", "Common Area", "Back of House", "Storage", "Other"],
    "Warehouse": ["Warehouse Floor", "Racking Area", "Office Area", "Loading Bay", "Mezzanine", "Yard", "Waste Area", "Other"],
    "Other": ["Area 1", "Area 2", "Room", "External Area", "Internal Area", "Other"]
}
 
room_options = room_options_map[project_type]
 
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
        st.button(f"➕ Add Work Item - {room_name}", key=f"add_item_{room_key}", on_click=add_work_item, args=(room_idx,), use_container_width=True)
    with a2:
        st.button(f"➖ Remove Last Item - {room_name}", key=f"remove_item_{room_key}", on_click=remove_work_item, args=(room_idx,), use_container_width=True)
 
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
 
        measurement_method = st.radio("Measurement Method", measurement_options, horizontal=True, key=f"measure_{item_key}")
 
        if measurement_method == "Wall Calculator (Total Length x Height)":
            w1, w2 = st.columns(2)
            with w1:
                total_wall_length = st.number_input("Total Wall Length (lm)", min_value=0.0, value=1.0, step=0.5, key=f"wall_length_{item_key}")
            with w2:
                wall_height = st.number_input("Wall Height (m)", min_value=0.0, value=2.7, step=0.1, key=f"wall_height_{item_key}")
 
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
                length = st.number_input("Length (m)", min_value=0.0, value=1.0, step=0.5, key=f"length_{item_key}")
            with l2:
                width = st.number_input("Width (m)", min_value=0.0, value=1.0, step=0.5, key=f"width_{item_key}")
            quantity = length * width
            measurement_note = f"Area used for pricing: {quantity:.2f} m²."
            st.info(measurement_note)
 
        else:
            quantity = st.number_input(f"Manual Quantity ({unit})", min_value=0.0, value=1.0, step=1.0, key=f"qty_{item_key}")
            measurement_note = f"Manual quantity used: {quantity:.2f} {unit}."
 
        p1, p2, p3 = st.columns([1, 1, 2])
        with p1:
            your_rate = st.number_input(f"Your Rate per {unit}", min_value=0.0, value=float(row["market_recommended"]), step=5.0, key=f"rate_{item_key}")
 
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
        pct = st.number_input(f"{name} %", min_value=0.0, max_value=100.0, value=default_pct, step=1.0, disabled=not checked, key=f"pct_{name}")
        if checked:
            selected_multiplier_total += pct
            multiplier_rows.append({"Condition": name, "Percent": pct})
 
use_custom = st.checkbox("Custom Extra %", value=False)
custom_pct = st.number_input("Custom Difficulty / Risk %", min_value=0.0, max_value=100.0, value=0.0, step=1.0, disabled=not use_custom)
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
 
def sale_price_from_margin(cost, margin_pct):
    if margin_pct <= 0 or margin_pct >= 100:
        return cost
    return cost / (1 - margin_pct / 100)
 
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
            quote_id = f"Q-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
            now = datetime.now().strftime("%d/%m/%Y %H:%M")
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
                    money(minimum_price),
                    money(recommended_price),
                    money(premium_price),
                    money(recommended_gst),
                    money(recommended_inc_gst),
                    str(row["Notes"]),
                    int(labourers),
                    round(float(estimated_days), 2),
                    int(rounded_days)
                ])
 
            ws.append_rows(rows, value_input_option="RAW")
            st.success(f"✅ Costing saved successfully. Quote ID: {quote_id}")
 
        except Exception as e:
            st.error("Could not save costing to Google Sheets.")
            st.exception(e)
