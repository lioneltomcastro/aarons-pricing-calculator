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

st.set_page_config(
    page_title="Aaron's Pricing Calculator",
    page_icon="🧾",
    layout="wide"
)

# =========================================================
# SIDEBAR
# =========================================================

page = st.sidebar.radio(
    "Navigation",
    [
        "New Costing",
        "Saved Costings"
    ]
)

st.title("🧾 Aaron's Pricing Calculator")
st.caption("Demolition, strip-out, flooring and rubbish removal cost calculator.")

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
        ws = spreadsheet.worksheet("Costing_Records")

    except:
        ws = spreadsheet.add_worksheet(
            title="Costing_Records",
            rows="5000",
            cols="50"
        )

        ws.append_row([
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
            "notes"
        ])

    return ws

# =========================================================
# LOAD SAVED COSTINGS
# =========================================================

def load_saved_costings():

    ws = connect_google_sheet()

    records = ws.get_all_records()

    if not records:
        return pd.DataFrame()

    return pd.DataFrame(records)

# =========================================================
# LOAD RATES
# =========================================================

@st.cache_data(ttl=300)
def load_rates(url):

    df = pd.read_csv(url)

    df.columns = [c.strip().lower() for c in df.columns]

    required = [
        "item",
        "category",
        "unit",
        "market_min",
        "market_recommended",
        "market_high",
        "productivity_per_day",
        "notes"
    ]

    missing = [c for c in required if c not in df.columns]

    if missing:
        st.error(f"Missing columns in Rates sheet: {missing}")
        st.stop()

    for c in [
        "market_min",
        "market_recommended",
        "market_high",
        "productivity_per_day"
    ]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    for c in ["item", "category", "unit", "notes"]:
        df[c] = df[c].astype(str).str.strip()

    return df

rates_df = load_rates(SHEET_CSV_URL)

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
        saved_df = saved_df[
            saved_df["client_name"].astype(str).str.contains(
                client_filter,
                case=False,
                na=False
            )
        ]

    if project_filter:
        saved_df = saved_df[
            saved_df["project_address"].astype(str).str.contains(
                project_filter,
                case=False,
                na=False
            )
        ]

    if quote_filter:
        saved_df = saved_df[
            saved_df["quote_id"].astype(str).str.contains(
                quote_filter,
                case=False,
                na=False
            )
        ]

    summary_df = saved_df.groupby(
        [
            "quote_id",
            "date",
            "client_name",
            "project_address"
        ],
        as_index=False
    ).agg({
        "final_total": "max"
    })

    summary_df = summary_df.sort_values(
        by="date",
        ascending=False
    )

    st.markdown("## Costing Records")

    st.dataframe(summary_df, use_container_width=True)

    st.markdown("## Open Costing")

    selected_quote = st.selectbox(
        "Select Quote ID",
        summary_df["quote_id"].unique()
    )

    selected_df = saved_df[
        saved_df["quote_id"] == selected_quote
    ]

    if not selected_df.empty:

        first_row = selected_df.iloc[0]

        d1, d2, d3 = st.columns(3)

        d1.metric(
            "Client",
            first_row["client_name"]
        )

        d2.metric(
            "Project",
            first_row["project_address"]
        )

        d3.metric(
            "Final Total",
            f"${float(first_row['final_total']):,.2f}"
        )

        st.markdown("### Work Item Breakdown")

        st.dataframe(selected_df, use_container_width=True)

    st.stop()

# =========================================================
# SESSION STATE
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
# PROJECT DETAILS
# =========================================================

st.subheader("1. Project Details")

c1, c2, c3 = st.columns(3)

with c1:

    client_name = st.text_input("Client Name")

    project_address = st.text_input("Project Address")

with c2:

    project_type = st.selectbox(
        "Project Type",
        [
            "Residential",
            "Commercial",
            "Industrial",
            "Shopping Centre",
            "Warehouse",
            "Other"
        ]
    )

    working_hours = st.selectbox(
        "Working Hours",
        [
            "Business Hours",
            "After Hours",
            "Weekend",
            "Day & Night",
            "To be confirmed"
        ]
    )

with c3:

    floor_level = st.number_input(
        "Floor Level",
        min_value=0,
        value=0,
        step=1
    )

    gst_rate = st.number_input(
        "GST %",
        min_value=0.0,
        max_value=20.0,
        value=10.0,
        step=0.5
    )

st.divider()

# =========================================================
# ROOM OPTIONS
# =========================================================

room_options_map = {

    "Residential": [
        "Kitchen",
        "Bathroom",
        "Laundry",
        "Bedroom",
        "Living Room",
        "Dining Room",
        "Hallway",
        "Garage",
        "Backyard",
        "Other"
    ],

    "Commercial": [
        "Tenancy Area",
        "Office",
        "Reception",
        "Meeting Room",
        "Kitchenette",
        "Bathroom",
        "Storage",
        "Shopfront",
        "Common Area",
        "Other"
    ],

    "Industrial": [
        "Warehouse Floor",
        "Racking Area",
        "Loading Bay",
        "Office Area",
        "Mezzanine",
        "Yard",
        "Concrete Area",
        "Storage Area",
        "Waste Area",
        "Other"
    ],

    "Shopping Centre": [
        "Tenancy Area",
        "Food Court",
        "Shopfront",
        "Kitchen",
        "Coolroom",
        "Common Area",
        "Back of House",
        "Storage",
        "Other"
    ],

    "Warehouse": [
        "Warehouse Floor",
        "Racking Area",
        "Office Area",
        "Loading Bay",
        "Mezzanine",
        "Yard",
        "Waste Area",
        "Other"
    ],

    "Other": [
        "Area 1",
        "Area 2",
        "Room",
        "External Area",
        "Internal Area",
        "Other"
    ]
}

room_options = room_options_map[project_type]

# =========================================================
# ROOMS / WORK ITEMS
# =========================================================

st.subheader("2. Rooms / Areas & Work Items")

b1, b2, b3 = st.columns([1, 1, 4])

with b1:
    st.button(
        "➕ Add Room / Area",
        on_click=add_room,
        use_container_width=True
    )

with b2:
    st.button(
        "➖ Remove Last Room",
        on_click=remove_room,
        use_container_width=True
    )

with b3:
    st.button(
        "🔄 Reset All Rooms",
        on_click=reset_all
    )

items = []

categories = ["All"] + sorted(
    rates_df["category"].dropna().unique().tolist()
)

for room_idx in range(st.session_state.room_count):

    room_key = f"{st.session_state.reset_counter}_{room_idx}"

    st.markdown(f"## Room / Area {room_idx + 1}")

    r1, r2 = st.columns(2)

    with r1:

        selected_room = st.selectbox(
            "Room / Area Type",
            room_options,
            key=f"room_type_{room_key}"
        )

    with r2:

        custom_room_name = st.text_input(
            "Custom Name (Optional)",
            key=f"custom_room_{room_key}"
        )

    room_name = (
        custom_room_name.strip()
        if custom_room_name.strip()
        else selected_room
    )

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

            category_filter = st.selectbox(
                "Category",
                categories,
                key=f"category_{item_key}"
            )

        filtered_df = (
            rates_df
            if category_filter == "All"
            else rates_df[
                rates_df["category"] == category_filter
            ]
        )

        with x2:

            selected_item = st.selectbox(
                "Work Type",
                filtered_df["item"].tolist(),
                key=f"work_type_{item_key}"
            )

        row = rates_df[
            rates_df["item"] == selected_item
        ].iloc[0]

        unit = str(row["unit"]).strip()

        unit_lower = unit.lower()

        item_lower = selected_item.lower()

        # =====================================================
        # MEASUREMENT
        # =====================================================

        st.markdown("#### Measurement")

        measurement_options = ["Manual Quantity"]

        if "wall" in item_lower or "partition" in item_lower:
            measurement_options.append(
                "Wall Calculator (Total Length x Height)"
            )

        if unit_lower in ["m2", "sqm", "sq m", "m²"]:
            measurement_options.append(
                "Area Calculator (Length x Width)"
            )

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

            if unit_lower in [
                "lm",
                "linear metre",
                "linear meter"
            ]:
                quantity = total_wall_length
            else:
                quantity = wall_area

            st.info(
                f"Calculated area: {wall_area:.2f} m²"
            )

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

            st.info(
                f"Calculated area: {quantity:.2f} m²"
            )

        else:

            quantity = st.number_input(
                f"Manual Quantity ({unit})",
                min_value=0.0,
                value=1.0,
                step=1.0,
                key=f"qty_{item_key}"
            )

        # =====================================================
        # RATE
        # =====================================================

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
            st.metric(
                "Item Total",
                f"${item_total:,.2f}"
            )

        with p3:
            st.caption(
                f"Quantity used: {quantity:.2f} {unit}"
            )

        market_min = float(row["market_min"])
        market_rec = float(row["market_recommended"])
        market_high = float(row["market_high"])
        productivity = float(row["productivity_per_day"])

        if your_rate < market_min:

            st.error(
                f"🚨 Below market minimum: "
                f"${market_min:,.2f}/{unit}"
            )

        elif market_min <= your_rate < market_rec:

            st.warning(
                f"⚠️ Below recommended: "
                f"${market_rec:,.2f}/{unit}"
            )

        elif market_rec <= your_rate <= market_high:

            st.success("✅ Within normal market range.")

        else:

            st.info(
                f"💎 Premium pricing. "
                f"Market high: ${market_high:,.2f}/{unit}"
            )

        worker_days = (
            quantity / productivity
            if productivity > 0
            else 0
        )

        notes = st.text_input(
            "Item Notes",
            value=str(row["notes"]),
            key=f"notes_{item_key}"
        )

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

# =========================================================
# TOTALS
# =========================================================

items_df = pd.DataFrame(items)

base_total = (
    float(items_df["Total"].sum())
    if not items_df.empty
    else 0
)

total_worker_days = (
    float(items_df["Worker Days"].sum())
    if not items_df.empty
    else 0
)

# =========================================================
# LABOUR
# =========================================================

st.subheader("3. Labour & Estimated Duration")

l1, l2, l3, l4 = st.columns(4)

with l1:
    labourers = st.number_input(
        "Number of Labourers",
        min_value=1,
        value=2,
        step=1
    )

with l2:
    labourer_day_rate = st.number_input(
        "Labourer Cost / Day",
        min_value=0.0,
        value=400.0,
        step=25.0
    )

with l3:
    supervisor_days = st.number_input(
        "Supervisor Days",
        min_value=0.0,
        value=0.0,
        step=0.5
    )

with l4:
    supervisor_day_rate = st.number_input(
        "Supervisor Cost / Day",
        min_value=0.0,
        value=550.0,
        step=25.0
    )

estimated_days = (
    total_worker_days / labourers
    if labourers > 0
    else 0
)

rounded_days = (
    ceil(estimated_days)
    if estimated_days > 0
    else 0
)

labour_cost = (
    rounded_days
    * labourers
    * labourer_day_rate
)

supervisor_cost = (
    supervisor_days
    * supervisor_day_rate
)

st.info(
    f"Estimated labour duration: "
    f"{estimated_days:.2f} days "
    f"({rounded_days} rounded day/s)"
)

# =========================================================
# EXTRA COSTS
# =========================================================

st.subheader("4. Extra Costs")

e1, e2, e3, e4 = st.columns(4)

with e1:
    equipment_cost = st.number_input(
        "Equipment Cost",
        min_value=0.0,
        value=0.0,
        step=100.0
    )

with e2:
    disposal_cost = st.number_input(
        "Disposal / Tip Fees",
        min_value=0.0,
        value=0.0,
        step=100.0
    )

with e3:
    truck_cost = st.number_input(
        "Truck / Transport Cost",
        min_value=0.0,
        value=0.0,
        step=100.0
    )

with e4:
    other_cost = st.number_input(
        "Other Cost",
        min_value=0.0,
        value=0.0,
        step=100.0
    )

st.divider()

# =========================================================
# MULTIPLIERS
# =========================================================

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

cols = st.columns(3)

for idx, (name, default_pct) in enumerate(
    default_multipliers.items()
):

    with cols[idx % 3]:

        checked = st.checkbox(
            name,
            value=False,
            key=f"check_{name}"
        )

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

# =========================================================
# FINAL PRICING
# =========================================================

st.subheader("6. Margin & Final Price")

m1, m2, m3 = st.columns(3)

with m1:
    minimum_margin = st.number_input(
        "Minimum Margin %",
        min_value=0.0,
        max_value=80.0,
        value=20.0,
        step=1.0
    )

with m2:
    recommended_margin = st.number_input(
        "Recommended Margin %",
        min_value=0.0,
        max_value=80.0,
        value=30.0,
        step=1.0
    )

with m3:
    premium_margin = st.number_input(
        "Premium Margin %",
        min_value=0.0,
        max_value=80.0,
        value=40.0,
        step=1.0
    )

direct_cost = (
    base_total
    + labour_cost
    + supervisor_cost
    + equipment_cost
    + disposal_cost
    + truck_cost
    + other_cost
)

difficulty_allowance = (
    direct_cost
    * (selected_multiplier_total / 100)
)

cost_with_difficulty = (
    direct_cost
    + difficulty_allowance
)

def apply_margin(cost, margin):

    if margin <= 0 or margin >= 100:
        return cost

    return cost / (1 - margin / 100)

minimum_price = apply_margin(
    cost_with_difficulty,
    minimum_margin
)

recommended_price = apply_margin(
    cost_with_difficulty,
    recommended_margin
)

premium_price = apply_margin(
    cost_with_difficulty,
    premium_margin
)

recommended_gst = (
    recommended_price
    * (gst_rate / 100)
)

recommended_inc_gst = (
    recommended_price
    + recommended_gst
)

# =========================================================
# SUMMARY
# =========================================================

st.divider()

st.subheader("7. Pricing Summary")

s1, s2, s3, s4 = st.columns(4)

s1.metric(
    "Base Work Items",
    f"${base_total:,.2f}"
)

s2.metric(
    "Labour Cost",
    f"${labour_cost:,.2f}"
)

s3.metric(
    "Difficulty %",
    f"{selected_multiplier_total:.1f}%"
)

s4.metric(
    "Difficulty Allowance",
    f"${difficulty_allowance:,.2f}"
)

s5, s6, s7, s8 = st.columns(4)

s5.metric(
    "Direct Cost",
    f"${direct_cost:,.2f}"
)

s6.metric(
    "Cost + Difficulty",
    f"${cost_with_difficulty:,.2f}"
)

s7.metric(
    "Recommended + GST",
    f"${recommended_price:,.2f}"
)

s8.metric(
    "Total Inc GST",
    f"${recommended_inc_gst:,.2f}"
)

st.markdown("### Suggested Prices")

p1, p2, p3 = st.columns(3)

p1.success(
    f"Minimum Safe Price: "
    f"${minimum_price:,.2f} + GST"
)

p2.info(
    f"Recommended Price: "
    f"${recommended_price:,.2f} + GST"
)

p3.warning(
    f"Premium Price: "
    f"${premium_price:,.2f} + GST"
)

# =========================================================
# BREAKDOWN
# =========================================================

st.markdown("### Work Item Breakdown")

if not items_df.empty:
    st.dataframe(items_df, use_container_width=True)

# =========================================================
# SAVE
# =========================================================

st.markdown("## Save Costing")

if st.button("💾 Save Costing to Google Sheets"):

    if items_df.empty:

        st.warning("No work items to save.")

    else:

        try:

            ws = connect_google_sheet()

            quote_id = (
                f"Q-"
                f"{datetime.now().strftime('%Y%m%d-%H%M%S')}"
            )

            now = datetime.now().strftime(
                "%d/%m/%Y %H:%M"
            )

            rows = []

            for _, row in items_df.iterrows():

                rows.append([
                    quote_id,
                    now,
                    client_name,
                    project_address,
                    project_type,
                    working_hours,
                    floor_level,
                    row["Room / Area"],
                    row["Item"],
                    row["Category"],
                    row["Unit"],
                    row["Measurement Method"],
                    row["Quantity Used"],
                    row["Your Rate"],
                    row["Total"],
                    row["Worker Days"],
                    equipment_cost,
                    disposal_cost,
                    truck_cost,
                    other_cost,
                    selected_multiplier_total,
                    difficulty_allowance,
                    labour_cost,
                    supervisor_cost,
                    direct_cost,
                    recommended_price,
                    recommended_gst,
                    recommended_inc_gst,
                    row["Notes"]
                ])

            ws.append_rows(
                rows,
                value_input_option="USER_ENTERED"
            )

            st.success(
                f"✅ Costing saved successfully. "
                f"Quote ID: {quote_id}"
            )

        except Exception as e:

            st.error(
                "Could not save costing to Google Sheets."
            )

            st.exception(e)
