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
st.caption("Cost calculator for demolition, strip-out, flooring and rubbish removal works.")

@st.cache_data(ttl=300)
def load_rates(url: str) -> pd.DataFrame:
    df = pd.read_csv(url)
    df.columns = [c.strip().lower() for c in df.columns]

    required_cols = [
        "item",
        "category",
        "unit",
        "market_min",
        "market_recommended",
        "market_high",
        "productivity_per_day",
        "notes",
    ]

    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Missing columns in Google Sheet: {missing}")
        st.stop()

    numeric_cols = ["market_min", "market_recommended", "market_high", "productivity_per_day"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["item"] = df["item"].astype(str).str.strip()
    df["category"] = df["category"].astype(str).str.strip()
    df["unit"] = df["unit"].astype(str).str.strip()
    df["notes"] = df["notes"].astype(str).fillna("")

    return df

try:
    rates_df = load_rates(SHEET_CSV_URL)
except Exception as e:
    st.error("Could not load the Google Sheet CSV. Please check the published CSV link.")
    st.exception(e)
    st.stop()

if "item_count" not in st.session_state:
    st.session_state.item_count = 1

if "reset_counter" not in st.session_state:
    st.session_state.reset_counter = 0

def add_item():
    st.session_state.item_count += 1

def remove_item():
    if st.session_state.item_count > 1:
        st.session_state.item_count -= 1

def reset_items():
    st.session_state.item_count = 1
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

st.subheader("2. Work Items")

btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 4])
with btn_col1:
    st.button("➕ Add Work Item", on_click=add_item, use_container_width=True)
with btn_col2:
    st.button("➖ Remove Last", on_click=remove_item, use_container_width=True)
with btn_col3:
    st.button("🔄 Reset Items", on_click=reset_items)

items = []
categories = ["All"] + sorted(rates_df["category"].dropna().unique().tolist())

for i in range(st.session_state.item_count):
    st.markdown(f"### Item {i + 1}")

    row_key = f"{st.session_state.reset_counter}_{i}"

    c1, c2, c3, c4, c5 = st.columns([1.3, 2, 0.8, 0.9, 0.9])

    with c1:
        category_filter = st.selectbox(
            "Category",
            categories,
            key=f"category_{row_key}"
        )

    filtered_df = rates_df if category_filter == "All" else rates_df[rates_df["category"] == category_filter]
    item_options = filtered_df["item"].tolist()

    with c2:
        selected_item = st.selectbox(
            "Work Type",
            item_options,
            key=f"item_{row_key}"
        )

    selected_row = rates_df[rates_df["item"] == selected_item].iloc[0]

    with c3:
        quantity = st.number_input(
            f"Qty ({selected_row['unit']})",
            min_value=0.0,
            value=1.0,
            step=1.0,
            key=f"qty_{row_key}"
        )

    with c4:
        default_rate = float(selected_row["market_recommended"])
        your_rate = st.number_input(
            "Your Rate",
            min_value=0.0,
            value=default_rate,
            step=5.0,
            key=f"rate_{row_key}"
        )

    with c5:
        item_total = quantity * your_rate
        st.metric("Item Total", f"${item_total:,.2f}")

    market_min = float(selected_row["market_min"])
    market_rec = float(selected_row["market_recommended"])
    market_high = float(selected_row["market_high"])
    productivity = float(selected_row["productivity_per_day"])

    if your_rate < market_min:
        st.error(f"🚨 Your rate is below market minimum. Min: ${market_min:,.2f}/{selected_row['unit']}")
    elif market_min <= your_rate < market_rec:
        st.warning(f"⚠️ Your rate is below recommended. Recommended: ${market_rec:,.2f}/{selected_row['unit']}")
    elif market_rec <= your_rate <= market_high:
        st.success(f"✅ Your rate is within the normal market range. Recommended: ${market_rec:,.2f}/{selected_row['unit']}")
    else:
        st.info(f"💎 Premium rate. Market high reference: ${market_high:,.2f}/{selected_row['unit']}")

    ref_col1, ref_col2, ref_col3, ref_col4 = st.columns(4)
    ref_col1.caption(f"Market Min: ${market_min:,.2f}/{selected_row['unit']}")
    ref_col2.caption(f"Recommended: ${market_rec:,.2f}/{selected_row['unit']}")
    ref_col3.caption(f"Market High: ${market_high:,.2f}/{selected_row['unit']}")
    ref_col4.caption(f"Productivity: {productivity:g} {selected_row['unit']}/worker/day")

    notes = st.text_input(
        "Item Notes",
        value=str(selected_row["notes"]) if str(selected_row["notes"]).lower() != "nan" else "",
        key=f"notes_{row_key}"
    )

    worker_days = quantity / productivity if productivity > 0 else 0

    items.append({
        "Item": selected_item,
        "Category": selected_row["category"],
        "Unit": selected_row["unit"],
        "Quantity": quantity,
        "Your Rate": your_rate,
        "Total": item_total,
        "Market Min": market_min,
        "Recommended": market_rec,
        "Market High": market_high,
        "Productivity": productivity,
        "Worker Days": worker_days,
        "Notes": notes
    })

    st.divider()

items_df = pd.DataFrame(items)
base_total = float(items_df["Total"].sum()) if not items_df.empty else 0
total_worker_days = float(items_df["Worker Days"].sum()) if not items_df.empty else 0

st.subheader("3. Labour & Estimated Duration")

lc1, lc2, lc3, lc4 = st.columns(4)

with lc1:
    labourers = st.number_input("Number of Labourers", min_value=1, value=2, step=1)

with lc2:
    labourer_day_rate = st.number_input("Labourer Cost / Day", min_value=0.0, value=400.0, step=25.0)

with lc3:
    supervisor_days = st.number_input("Supervisor Days", min_value=0.0, value=0.0, step=0.5)

with lc4:
    supervisor_day_rate = st.number_input("Supervisor Cost / Day", min_value=0.0, value=550.0, step=25.0)

estimated_days = total_worker_days / labourers if labourers > 0 else 0
rounded_days = ceil(estimated_days) if estimated_days > 0 else 0

labour_cost = rounded_days * labourers * labourer_day_rate
supervisor_cost = supervisor_days * supervisor_day_rate

st.info(f"Estimated labour duration: **{estimated_days:.2f} days** based on productivity data. Rounded program allowance: **{rounded_days} day(s)**.")

st.subheader("4. Extra Costs")

ec1, ec2, ec3, ec4 = st.columns(4)

with ec1:
    equipment_cost = st.number_input("Equipment Cost", min_value=0.0, value=0.0, step=100.0)

with ec2:
    disposal_cost = st.number_input("Disposal / Tip Fees", min_value=0.0, value=0.0, step=100.0)

with ec3:
    truck_cost = st.number_input("Truck / Transport Cost", min_value=0.0, value=0.0, step=100.0)

with ec4:
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

cols = st.columns(3)
for idx, (name, default_pct) in enumerate(default_multipliers.items()):
    with cols[idx % 3]:
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

mc1, mc2, mc3 = st.columns(3)
with mc1:
    minimum_margin = st.number_input("Minimum Margin %", min_value=0.0, max_value=80.0, value=20.0, step=1.0)
with mc2:
    recommended_margin = st.number_input("Recommended Margin %", min_value=0.0, max_value=80.0, value=30.0, step=1.0)
with mc3:
    premium_margin = st.number_input("Premium Margin %", min_value=0.0, max_value=80.0, value=40.0, step=1.0)

direct_cost = base_total + labour_cost + supervisor_cost + equipment_cost + disposal_cost + truck_cost + other_cost
difficulty_allowance = direct_cost * (selected_multiplier_total / 100)
cost_with_difficulty = direct_cost + difficulty_allowance

def sale_price_from_margin(cost: float, margin_pct: float) -> float:
    if margin_pct >= 100:
        return cost
    return cost / (1 - margin_pct / 100) if margin_pct > 0 else cost

minimum_price = sale_price_from_margin(cost_with_difficulty, minimum_margin)
recommended_price = sale_price_from_margin(cost_with_difficulty, recommended_margin)
premium_price = sale_price_from_margin(cost_with_difficulty, premium_margin)

recommended_gst = recommended_price * (gst_rate / 100)
recommended_inc_gst = recommended_price + recommended_gst

st.divider()
st.subheader("7. Pricing Summary")

r1, r2, r3, r4 = st.columns(4)
r1.metric("Base Work Items", f"${base_total:,.2f}")
r2.metric("Labour Cost", f"${labour_cost:,.2f}")
r3.metric("Difficulty %", f"{selected_multiplier_total:.1f}%")
r4.metric("Difficulty Allowance", f"${difficulty_allowance:,.2f}")

r5, r6, r7, r8 = st.columns(4)
r5.metric("Direct Cost", f"${direct_cost:,.2f}")
r6.metric("Cost + Difficulty", f"${cost_with_difficulty:,.2f}")
r7.metric("Recommended + GST", f"${recommended_price:,.2f}")
r8.metric("Total Inc GST", f"${recommended_inc_gst:,.2f}")

st.markdown("### Suggested Prices")
p1, p2, p3 = st.columns(3)
p1.success(f"Minimum Safe Price: **${minimum_price:,.2f} + GST**")
p2.info(f"Recommended Price: **${recommended_price:,.2f} + GST**")
p3.warning(f"Premium Price: **${premium_price:,.2f} + GST**")

st.markdown("### Work Item Breakdown")
display_df = items_df.copy()
if not display_df.empty:
    money_cols = ["Your Rate", "Total", "Market Min", "Recommended", "Market High"]
    for col in money_cols:
        display_df[col] = display_df[col].map(lambda x: f"${x:,.2f}")
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
