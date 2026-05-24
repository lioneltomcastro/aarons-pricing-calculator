import streamlit as st
import pandas as pd
from math import ceil

# =========================================================
# Aaron's Pricing Calculator
# =========================================================

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQi2XaKTk3vZnIWejQCiCVNqLwJRcDKYThJCKHOH4iPA_JdDQFEUTcKq5BYRDtbAFDYcu6gWnqQgH2/pub?output=csv"

st.set_page_config(
    page_title="Aaron's Pricing Calculator",
    page_icon="🧾",
    layout="wide"
)

st.title("🧾 Aaron's Pricing Calculator")
st.caption("Demolition, strip-out, flooring and rubbish removal cost calculator.")

# =========================================================
# LOAD GOOGLE SHEET
# =========================================================

@st.cache_data(ttl=300)
def load_rates(url):
    df = pd.read_csv(url)

    df.columns = [c.strip().lower() for c in df.columns]

    numeric_cols = [
        "market_min",
        "market_recommended",
        "market_high",
        "productivity_per_day"
    ]

    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df

rates_df = load_rates(SHEET_CSV_URL)

# =========================================================
# SESSION STATE
# =========================================================

if "room_count" not in st.session_state:
    st.session_state.room_count = 1

if "room_items" not in st.session_state:
    st.session_state.room_items = {0: 1}

def add_room():
    st.session_state.room_count += 1
    st.session_state.room_items[st.session_state.room_count - 1] = 1

def remove_room():
    if st.session_state.room_count > 1:
        last_room = st.session_state.room_count - 1
        del st.session_state.room_items[last_room]
        st.session_state.room_count -= 1

def add_item(room_idx):
    st.session_state.room_items[room_idx] += 1

def remove_item(room_idx):
    if st.session_state.room_items[room_idx] > 1:
        st.session_state.room_items[room_idx] -= 1

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
        ["Residential", "Commercial", "Industrial"]
    )

    working_hours = st.selectbox(
        "Working Hours",
        ["Business Hours", "After Hours", "Weekend"]
    )

with c3:
    floor_level = st.number_input("Floor Level", min_value=0, value=0)
    gst_rate = st.number_input("GST %", min_value=0.0, value=10.0)

st.divider()

# =========================================================
# ROOM OPTIONS
# =========================================================

if project_type == "Residential":
    room_options = [
        "Kitchen",
        "Bathroom",
        "Laundry",
        "Bedroom",
        "Living Room",
        "Garage",
        "Other"
    ]

elif project_type == "Commercial":
    room_options = [
        "Office",
        "Reception",
        "Meeting Room",
        "Tenancy Area",
        "Shopfront",
        "Storage",
        "Other"
    ]

else:
    room_options = [
        "Warehouse Floor",
        "Office Area",
        "Loading Bay",
        "Yard",
        "Storage Area",
        "Other"
    ]

# =========================================================
# WORK ITEMS
# =========================================================

st.subheader("2. Rooms / Areas & Work Items")

b1, b2 = st.columns(2)

with b1:
    st.button("➕ Add Room / Area", on_click=add_room)

with b2:
    st.button("➖ Remove Last Room", on_click=remove_room)

items = []

categories = ["All"] + sorted(rates_df["category"].dropna().unique())

for room_idx in range(st.session_state.room_count):

    st.markdown(f"## Room / Area {room_idx + 1}")

    rc1, rc2 = st.columns(2)

    with rc1:
        selected_room = st.selectbox(
            "Room Type",
            room_options,
            key=f"room_type_{room_idx}"
        )

    with rc2:
        custom_room = st.text_input(
            "Custom Room Name",
            key=f"custom_room_{room_idx}"
        )

    room_name = custom_room if custom_room else selected_room

    ib1, ib2 = st.columns(2)

    with ib1:
        st.button(
            f"➕ Add Item - {room_name}",
            key=f"add_item_{room_idx}",
            on_click=add_item,
            args=(room_idx,)
        )

    with ib2:
        st.button(
            f"➖ Remove Item - {room_name}",
            key=f"remove_item_{room_idx}",
            on_click=remove_item,
            args=(room_idx,)
        )

    for item_idx in range(st.session_state.room_items[room_idx]):

        st.markdown(f"### {room_name} - Item {item_idx + 1}")

        key = f"{room_idx}_{item_idx}"

        c1, c2 = st.columns([1, 2])

        with c1:
            category_filter = st.selectbox(
                "Category",
                categories,
                key=f"category_{key}"
            )

        filtered_df = rates_df if category_filter == "All" else rates_df[rates_df["category"] == category_filter]

        item_options = filtered_df["item"].tolist()

        with c2:
            selected_item = st.selectbox(
                "Work Type",
                item_options,
                key=f"item_{key}"
            )

        selected_row = rates_df[rates_df["item"] == selected_item].iloc[0]

        selected_unit = str(selected_row["unit"]).lower()

        # =========================================================
        # MEASUREMENT
        # =========================================================

        st.markdown("#### Measurement")

        measurement_options = ["Manual Quantity"]

        item_name = selected_item.lower()

        if "wall" in item_name or "partition" in item_name:
            measurement_options.append("Wall Calculator")

        if selected_unit in ["m2", "sqm", "m²"]:
            measurement_options.append("Area Calculator")

        measurement_method = st.radio(
            "Measurement Method",
            measurement_options,
            horizontal=True,
            key=f"measurement_{key}"
        )

        calculated_qty = 0.0

        # =========================================================
        # WALL CALCULATOR
        # =========================================================

        if measurement_method == "Wall Calculator":

            wc1, wc2 = st.columns(2)

            with wc1:
                total_wall_length = st.number_input(
                    "Total Wall Length (lm)",
                    min_value=0.0,
                    value=1.0,
                    step=0.5,
                    key=f"wall_length_{key}"
                )

            with wc2:
                wall_height = st.number_input(
                    "Wall Height (m)",
                    min_value=0.0,
                    value=2.7,
                    step=0.1,
                    key=f"wall_height_{key}"
                )

            calculated_area = total_wall_length * wall_height

            if selected_unit == "lm":
                calculated_qty = total_wall_length
            else:
                calculated_qty = calculated_area

            st.info(f"Calculated Area: {calculated_area:.2f} m²")

        # =========================================================
        # AREA CALCULATOR
        # =========================================================

        elif measurement_method == "Area Calculator":

            ac1, ac2 = st.columns(2)

            with ac1:
                area_length = st.number_input(
                    "Length (m)",
                    min_value=0.0,
                    value=1.0,
                    step=0.5,
                    key=f"area_length_{key}"
                )

            with ac2:
                area_width = st.number_input(
                    "Width (m)",
                    min_value=0.0,
                    value=1.0,
                    step=0.5,
                    key=f"area_width_{key}"
                )

            calculated_qty = area_length * area_width

            st.info(f"Calculated Area: {calculated_qty:.2f} m²")

        # =========================================================
        # MANUAL QUANTITY
        # =========================================================

        else:

            calculated_qty = st.number_input(
                f"Quantity ({selected_row['unit']})",
                min_value=0.0,
                value=1.0,
                step=1.0,
                key=f"qty_{key}"
            )

        # =========================================================
        # RATE
        # =========================================================

        r1, r2 = st.columns(2)

        with r1:
            your_rate = st.number_input(
                f"Your Rate per {selected_row['unit']}",
                min_value=0.0,
                value=float(selected_row["market_recommended"]),
                step=5.0,
                key=f"rate_{key}"
            )

        item_total = calculated_qty * your_rate

        with r2:
            st.metric("Item Total", f"${item_total:,.2f}")

        market_min = float(selected_row["market_min"])
        market_rec = float(selected_row["market_recommended"])
        market_high = float(selected_row["market_high"])

        if your_rate < market_min:
            st.error(f"🚨 Your rate is below market minimum. Min: ${market_min:,.2f}/{selected_row['unit']}")

        elif market_min <= your_rate < market_rec:
            st.warning(f"⚠️ Your rate is below recommended. Recommended: ${market_rec:,.2f}/{selected_row['unit']}")

        elif market_rec <= your_rate <= market_high:
            st.success("✅ Your rate is within market range.")

        else:
            st.info("💎 Premium commercial pricing.")

        st.caption(f"Market Min: ${market_min:,.2f}/{selected_row['unit']}")
        st.caption(f"Recommended: ${market_rec:,.2f}/{selected_row['unit']}")
        st.caption(f"Market High: ${market_high:,.2f}/{selected_row['unit']}")

        items.append({
            "Room": room_name,
            "Item": selected_item,
            "Quantity": calculated_qty,
            "Rate": your_rate,
            "Total": item_total
        })

        st.divider()

# =========================================================
# SUMMARY
# =========================================================

items_df = pd.DataFrame(items)

base_total = items_df["Total"].sum() if not items_df.empty else 0

st.subheader("3. Pricing Summary")

st.metric("Total Price", f"${base_total:,.2f} + GST")

gst_total = base_total * (gst_rate / 100)

st.metric("Total Including GST", f"${base_total + gst_total:,.2f}")
