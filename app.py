import streamlit as st
import pandas as pd
from math import ceil

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQi2XaKTk3vZnIWejQCiCVNqLwJRcDKYThJCKHOH4iPA_JdDQFEUTcKq5BYRDtbAFDYcu6gWnqQgH2/pub?output=csv"

st.set_page_config(
    page_title="Aaron's Pricing Calculator",
    page_icon="🧾",
    layout="wide"
)

st.title("🧾 Aaron's Pricing Calculator")
st.caption("Demolition, strip-out, flooring and rubbish removal cost calculator.")

@st.cache_data(ttl=300)
def load_rates(url):
    df = pd.read_csv(url)
    df.columns = [c.strip().lower() for c in df.columns]

    required_cols = [
        "item", "category", "unit", "market_min",
        "market_recommended", "market_high",
        "productivity_per_day", "notes"
    ]

    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Missing columns in Google Sheet: {missing}")
        st.stop()

    for col in ["market_min", "market_recommended", "market_high", "productivity_per_day"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["item"] = df["item"].astype(str).str.strip()
    df["category"] = df["category"].astype(str).str.strip()
    df["unit"] = df["unit"].astype(str).str.strip()
    df["notes"] = df["notes"].astype(str).fillna("")

    return df

try:
    rates_df = load_rates(SHEET_CSV_URL)
except Exception as e:
    st.error("Could not load Google Sheet CSV.")
    st.exception(e)
    st.stop()

if "room_count" not in st.session_state:
    st.session_state.room_count = 1

if "room_items" not in st.session_state:
    st.session_state.room_items = {0: 1}

if "reset_counter" not in st.session_state:
    st.session_state.reset_counter = 0

def add_room():
    new_room = st.session_state.room_count
    st.session_state.room_count += 1
    st.session_state.room_items[new_room] = 1

def remove_room():
    if st.session_state.room_count > 1:
        last_room = st.session_state.room_count - 1
        st.session_state.room_items.pop(last_room, None)
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

st.subheader("1. Project Details")

col1, col2, col3 = st.columns(3)

with col1:
    client_name = st.text_input("Client Name", placeholder="e.g. John / Builder / Company")
    project_address = st.text_input("Project Address", placeholder="e.g. Reservoir VIC")

with col2:
    project_type = st.selectbox(
        "Project Type",
        ["Residential", "Commercial", "Industrial", "Shopping Centre", "Warehouse", "Other"]
    )

    working_hours = st.selectbox(
        "Working Hours",
        ["Business Hours", "After Hours", "Weekend", "Day & Night", "To be confirmed"]
    )

with col3:
    floor_level = st.number_input("Floor Level", min_value=0, value=0, step=1)
    gst_rate = st.number_input("GST %", min_value=0.0, max_value=20.0, value=10.0, step=0.5)

st.divider()

if project_type == "Residential":
    room_options = [
        "Kitchen", "Bathroom", "Laundry", "Bedroom", "Living Room",
        "Dining Room", "Hallway", "Garage", "Backyard", "Other"
    ]
elif project_type == "Commercial":
    room_options = [
        "Tenancy Area", "Office", "Reception", "Meeting Room", "Kitchenette",
        "Bathroom", "Storage", "Shopfront", "Common Area", "Other"
    ]
elif project_type == "Industrial":
    room_options = [
        "Warehouse Floor", "Racking Area", "Loading Bay", "Office Area",
        "Mezzanine", "Yard", "Concrete Area", "Storage Area", "Waste Area", "Other"
    ]
elif project_type == "Shopping Centre":
    room_options = [
        "Tenancy Area", "Food Court", "Shopfront", "Kitchen", "Coolroom",
        "Common Area", "Back of House", "Storage", "Other"
    ]
elif project_type == "Warehouse":
    room_options = [
        "Warehouse Floor", "Racking Area", "Office Area", "Loading Bay",
        "Mezzanine", "Yard", "Waste Area", "Other"
    ]
else:
    room_options = ["Area 1", "Area 2", "Room", "External Area", "Internal Area", "Other"]

st.subheader("2. Rooms / Areas & Work Items")

btn1, btn2, btn3 = st.columns([1, 1, 4])

with btn1:
    st.button("➕ Add Room / Area", on_click=add_room, use_container_width=True)

with btn2:
    st.button("➖ Remove Last Room", on_click=remove_room, use_container_width=True)

with btn3:
    st.button("🔄 Reset All Rooms", on_click=reset_all)

items = []
categories = ["All"] + sorted(rates_df["category"].dropna().unique().tolist())

for room_idx in range(st.session_state.room_count):
    room_key = f"{st.session_state.reset_counter}_{room_idx}"

    st.markdown(f"## Room / Area {room_idx + 1}")

    room_col1, room_col2 = st.columns(2)

    with room_col1:
        selected_room = st.selectbox(
            "Room / Area Type",
            room_options,
            key=f"room_type_{room_key}"
        )

    with room_col2:
        custom_room_name = st.text_input(
            "Custom Name (Optional)",
            placeholder="e.g. Bedroom 1 / Tenancy A / Warehouse Section 3",
            key=f"custom_room_{room_key}"
        )

    room_name = custom_room_name.strip() if custom_room_name.strip() else selected_room

    add_col1, add_col2 = st.columns(2)

    with add_col1:
        st.button(
            f"➕ Add Work Item - {room_name}",
            key=f"add_item_btn_{room_key}",
            on_click=add_work_item,
            args=(room_idx,),
            use_container_width=True
        )

    with add_col2:
        st.button(
            f"➖ Remove Last Item - {room_name}",
            key=f"remove_item_btn_{room_key}",
            on_click=remove_work_item,
            args=(room_idx,),
            use_container_width=True
        )

    for item_idx in range(st.session_state.room_items[room_idx]):
        item_key = f"{st.session_state.reset_counter}_{room_idx}_{item_idx}"

        st.markdown(f"### {room_name} - Item {item_idx + 1}")

        item_col1, item_col2 = st.columns([1.2, 2])

        with item_col1:
            category_filter = st.selectbox(
                "Category",
                categories,
                key=f"category_{item_key}"
            )

        filtered_df = rates_df if category_filter == "All" else rates_df[rates_df["category"] == category_filter]

        if filtered_df.empty:
            st.warning("No items found for this category.")
            continue

        item_options = filtered_df["item"].tolist()

        with item_col2:
            selected_item = st.selectbox(
                "Work Type",
                item_options,
                key=f"work_type_{item_key}"
            )

        selected_row = rates_df[rates_df["item"] == selected_item].iloc[0]
        selected_unit = str(selected_row["unit"]).strip()
        unit_lower = selected_unit.lower()
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
            key=f"measurement_method_{item_key}"
        )

        quantity_used = 0.0
        measurement_note = ""

        if measurement_method == "Wall Calculator (Total Length x Height)":
            wc1, wc2 = st.columns(2)

            with wc1:
                total_wall_length = st.number_input(
                    "Total Wall Length (lm)",
                    min_value=0.0,
                    value=1.0,
                    step=0.5,
                    key=f"wall_length_{item_key}"
                )

            with wc2:
                wall_height = st.number_input(
                    "Wall Height (m)",
                    min_value=0.0,
                    value=2.7,
                    step=0.1,
                    key=f"wall_height_{item_key}"
                )

            wall_area = total_wall_length * wall_height

            if unit_lower in ["lm", "linear metre", "linear meter"]:
                quantity_used = total_wall_length
                measurement_note = (
                    f"Calculated wall area: {wall_area:.2f} m². "
                    f"Pricing quantity used: {quantity_used:.2f} lm."
                )
            else:
                quantity_used = wall_area
                measurement_note = f"Calculated wall area used for pricing: {quantity_used:.2f} m²."

            st.info(measurement_note)

        elif measurement_method == "Area Calculator (Length x Width)":
            ac1, ac2 = st.columns(2)

            with ac1:
                area_length = st.number_input(
                    "Length (m)",
                    min_value=0.0,
                    value=1.0,
                    step=0.5,
                    key=f"area_length_{item_key}"
                )

            with ac2:
                area_width = st.number_input(
                    "Width (m)",
                    min_value=0.0,
                    value=1.0,
                    step=0.5,
                    key=f"area_width_{item_key}"
                )

            quantity_used = area_length * area_width
            measurement_note = f"Calculated area used for pricing: {quantity_used:.2f} m²."
            st.info(measurement_note)

        else:
            quantity_used = st.number_input(
                f"Manual Quantity ({selected_unit})",
                min_value=0.0,
                value=1.0,
                step=1.0,
                key=f"manual_qty_{item_key}"
            )
            measurement_note = f"Manual quantity used: {quantity_used:.2f} {selected_unit}."

        rate_col1, rate_col2, rate_col3 = st.columns([1, 1, 2])

        with rate_col1:
            your_rate = st.number_input(
                f"Your Rate per {selected_unit}",
                min_value=0.0,
                value=float(selected_row["market_recommended"]),
                step=5.0,
                key=f"your_rate_{item_key}"
            )

        item_total = quantity_used * your_rate

        with rate_col2:
            st.metric("Item Total", f"${item_total:,.2f}")

        with rate_col3:
            st.caption(f"Quantity used for pricing: {quantity_used:.2f} {selected_unit}")

        market_min = float(selected_row["market_min"])
        market_rec = float(selected_row["market_recommended"])
        market_high = float(selected_row["market_high"])
        productivity = float(selected_row["productivity_per_day"])

        if your_rate < market_min:
            st.error(f"🚨 Your rate is below market minimum. Min: ${market_min:,.2f}/{selected_unit}")
        elif market_min <= your_rate < market_rec:
            st.warning(f"⚠️ Your rate is below recommended. Recommended: ${market_rec:,.2f}/{selected_unit}")
        elif market_rec <= your_rate <= market_high:
            st.success("✅ Your rate is within normal market range.")
        else:
            st.info(f"💎 Premium commercial pricing. Market high: ${market_high:,.2f}/{selected_unit}")

        ref1, ref2, ref3, ref4 = st.columns(4)

        ref1.caption(f"Market Min: ${market_min:,.2f}/{selected_unit}")
        ref2.caption(f"Recommended: ${market_rec:,.2f}/{selected_unit}")
        ref3.caption(f"Market High: ${market_high:,.2f}/{selected_unit}")
        ref4.caption(f"Productivity: {productivity:g} {selected_unit}/worker/day")

        notes = st.text_input(
            "Item Notes",
            value=str(selected_row["notes"]) if str(selected_row["notes"]).lower() != "nan" else "",
            key=f"notes_{item_key}"
        )

        worker_days = quantity_used / productivity if productivity > 0 else 0

        items.append({
            "Room / Area": room_name,
            "Item": selected_item,
            "Category": selected_row["category"],
            "Unit": selected_unit,
            "Measurement Method": measurement_method,
            "Quantity Used": quantity_used,
            "Your Rate": your_rate,
            "Total": item_total,
            "Market Min": market_min,
            "Recommended": market_rec,
            "Market High": market_high,
            "Productivity": productivity,
            "Worker Days": worker_days,
            "Measurement Note": measurement_note,
            "Notes": notes
        })

        st.divider()

items_df = pd.DataFrame(items)
base_total = float(items_df["Total"].sum()) if not items_df.empty else 0
total_worker_days = float(items_df["Worker Days"].sum()) if not items_df.empty else 0

st.subheader("3. Labour & Estimated Duration")

lab_col1, lab_col2, lab_col3, lab_col4 = st.columns(4)

with lab_col1:
    labourers = st.number_input("Number of Labourers", min_value=1, value=2, step=1)

with lab_col2:
    labourer_day_rate = st.number_input("Labourer Cost / Day", min_value=0.0, value=400.0, step=25.0)

with lab_col3:
    supervisor_days = st.number_input("Supervisor Days", min_value=0.0, value=0.0, step=0.5)

with lab_col4:
    supervisor_day_rate = st.number_input("Supervisor Cost / Day", min_value=0.0, value=550.0, step=25.0)

estimated_days = total_worker_days / labourers if labourers > 0 else 0
rounded_days = ceil(estimated_days) if estimated_days > 0 else 0

labour_cost = rounded_days * labourers * labourer_day_rate
supervisor_cost = supervisor_days * supervisor_day_rate

st.info(
    f"Estimated labour duration: **{estimated_days:.2f} days** based on productivity data. "
    f"Rounded program allowance: **{rounded_days} day(s)**."
)

st.subheader("4. Extra Costs")

extra_col1, extra_col2, extra_col3, extra_col4 = st.columns(4)

with extra_col1:
    equipment_cost = st.number_input("Equipment Cost", min_value=0.0, value=0.0, step=100.0)

with extra_col2:
    disposal_cost = st.number_input("Disposal / Tip Fees", min_value=0.0, value=0.0, step=100.0)

with extra_col3:
    truck_cost = st.number_input("Truck / Transport Cost", min_value=0.0, value=0.0, step=100.0)

with extra_col4:
    other_cost = st.number_input("Other Cost", min_value=0.0, value=0.0, step=100.0)

st.divider()

st.subheader("5. Difficulty & Access Multipliers")
st.caption("Tick only the conditions that apply. You can edit each % manually.")

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
    "EWP Required": 20.0,
    "Traffic Management Required": 15.0,
    "Protection Works Required": 10.0,
}

selected_multiplier_total = 0.0
multiplier_rows = []

mult_cols = st.columns(3)

for idx, (name, default_pct) in enumerate(default_multipliers.items()):
    with mult_cols[idx % 3]:
        checked = st.checkbox(name, value=False, key=f"mult_check_{name}")
        pct = st.number_input(
            f"{name} %",
            min_value=0.0,
            max_value=100.0,
            value=default_pct,
            step=1.0,
            key=f"mult_pct_{name}",
            disabled=not checked
        )

        if checked:
            selected_multiplier_total += pct
            multiplier_rows.append({"Condition": name, "Percent": pct})

custom_col1, custom_col2 = st.columns([1, 3])

with custom_col1:
    use_custom = st.checkbox("Custom Extra %", value=False)

with custom_col2:
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

margin_col1, margin_col2, margin_col3 = st.columns(3)

with margin_col1:
    minimum_margin = st.number_input("Minimum Margin %", min_value=0.0, max_value=80.0, value=20.0, step=1.0)

with margin_col2:
    recommended_margin = st.number_input("Recommended Margin %", min_value=0.0, max_value=80.0, value=30.0, step=1.0)

with margin_col3:
    premium_margin = st.number_input("Premium Margin %", min_value=0.0, max_value=80.0, value=40.0, step=1.0)

direct_cost = base_total + labour_cost + supervisor_cost + equipment_cost + disposal_cost + truck_cost + other_cost
difficulty_allowance = direct_cost * (selected_multiplier_total / 100)
cost_with_difficulty = direct_cost + difficulty_allowance

def sale_price_from_margin(cost, margin_pct):
    if margin_pct >= 100:
        return cost
    if margin_pct <= 0:
        return cost
    return cost / (1 - margin_pct / 100)

minimum_price = sale_price_from_margin(cost_with_difficulty, minimum_margin)
recommended_price = sale_price_from_margin(cost_with_difficulty, recommended_margin)
premium_price = sale_price_from_margin(cost_with_difficulty, premium_margin)

recommended_gst = recommended_price * (gst_rate / 100)
recommended_inc_gst = recommended_price + recommended_gst

st.divider()
st.subheader("7. Pricing Summary")

sum_col1, sum_col2, sum_col3, sum_col4 = st.columns(4)

sum_col1.metric("Base Work Items", f"${base_total:,.2f}")
sum_col2.metric("Labour Cost", f"${labour_cost:,.2f}")
sum_col3.metric("Difficulty %", f"{selected_multiplier_total:.1f}%")
sum_col4.metric("Difficulty Allowance", f"${difficulty_allowance:,.2f}")

sum_col5, sum_col6, sum_col7, sum_col8 = st.columns(4)

sum_col5.metric("Direct Cost", f"${direct_cost:,.2f}")
sum_col6.metric("Cost + Difficulty", f"${cost_with_difficulty:,.2f}")
sum_col7.metric("Recommended + GST", f"${recommended_price:,.2f}")
sum_col8.metric("Total Inc GST", f"${recommended_inc_gst:,.2f}")

st.markdown("### Suggested Prices")

price_col1, price_col2, price_col3 = st.columns(3)

price_col1.success(f"Minimum Safe Price: **${minimum_price:,.2f} + GST**")
price_col2.info(f"Recommended Price: **${recommended_price:,.2f} + GST**")
price_col3.warning(f"Premium Price: **${premium_price:,.2f} + GST**")

st.markdown("### Work Item Breakdown")

display_df = items_df.copy()

if not display_df.empty:
    money_cols = ["Your Rate", "Total", "Market Min", "Recommended", "Market High"]

    for col in money_cols:
        display_df[col] = display_df[col].map(lambda x: f"${x:,.2f}")

    display_df["Quantity Used"] = display_df["Quantity Used"].map(lambda x: f"{x:.2f}")
    display_df["Worker Days"] = display_df["Worker Days"].map(lambda x: f"{x:.2f}")

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

Base Work Items: ${base_total:,.2f}
Labour Cost: ${labour_cost:,.2f}
Supervisor Cost: ${supervisor_cost:,.2f}
Equipment Cost: ${equipment_cost:,.2f}
Disposal / Tip Fees: ${disposal_cost:,.2f}
Truck / Transport Cost: ${truck_cost:,.2f}
Other Cost: ${other_cost:,.2f}
Difficulty Allowance ({selected_multiplier_total:.1f}%): ${difficulty_allowance:,.2f}

Minimum Safe Price: ${minimum_price:,.2f} + GST
Recommended Price: ${recommended_price:,.2f} + GST
Premium Price: ${premium_price:,.2f} + GST

Recommended Total Inc GST: ${recommended_inc_gst:,.2f}
Estimated Duration: {estimated_days:.2f} days ({rounded_days} rounded day/s)
"""

st.text_area("Copy this pricing summary", value=summary_text.strip(), height=320)
