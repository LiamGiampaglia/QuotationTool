from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
import streamlit as st
from docx import Document
import tempfile
from datetime import datetime
import openpyxl
import io


material_code_options = [
    # POWER CONSULTANCY
    "FSC11 EcoConsult Essential / MPS Walkthrough",
    "FSC719 Eco Audit Advanced Plus / MPS Enterprise",
    "FSC9002 Electrical Digital Twin",
    "FSC830 Re-Assessment Audit",
    "FSC612 EcoConsult Advanced",
    "FSC9007 Arc Flash_ETAP",
    "FSC9006 Protection Study_ETAP",
    "CON594 Fault Level Study",
    "CON607 Protection System Desig",
    "CON602 Earthing Design",
    "FSC5099 Relay Programming",
    "FSC5099 Pressure Rise Studies",
    "FSC5099 Network Design",
    "FSC5099 Substation Design",
    "FSC5099 Equipment Design",
    "SRV_ETAP_PQ_MOD Harmonic Surveys / Temp Monitoring",
    "FSC9011 ETAP Subscription",
    "SRV_PM_AUDIT_ADV EcoConsult Audit Power Monitoring audit",
    "SRV_PQ_AUDIT_ESS EcoConsult Audit PQ Essential",

    # PROCESS ELECTRIFICATION
    "PEPCONSULTING_0001 Process Electrification Project",

    # INDUSTRY & DT
    "GCRLIFECYCLECONSUL Industry Consulting (IDIBS)",
    "GCRDIGITRANSPRJCT Digital Transformation Project",

    # DATA CENTRE & COOLING
    "WCONSULTADV EcoConsult for Data Centers",

    # BMS
    "CON596 BMS Consultancy",

    # EV
    "EVS1AG eMobility Consultancy Audit",

    # ENERGY
    "CON200 Sust.Serv Smart Grid Elect.Audi",
    "CON201 Sust.Serv Smart Grid Elect.Design",
    "CON202 Sust.Serv EV Infrastructure Audit",
    "CON203 Sust.Serv Energy Efficiency Audit",
    "CON204 Sust.Serv Remote Energy Assessment",
    "CON205 Sust.Serv Metering Study",
    "CON206 Sust.Serv Energy Monitoring",

    # MICROGRID
    "SRVINAMGFEAS Microgrid Feasibility Study",
    "SRVINAMGDES Microgrid Design",

    # EXTRA
    "CON601 DON’T USE Consultancy Ancillary Offers",
    "CON595 Electrical Network Design"
]

def extract_rates(uploaded_file):

    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb["PRICING SHEET"]

    
    rates = {
        "office_cost": ws["E97"].value or 0,
        "site_cost": ws["E98"].value or 0,
    
        "office_day": ws["E101"].value or 0,
        "office_evening": ws["F101"].value or 0,
        "office_weekend": ws["G101"].value or 0,
    
        "site_day": ws["E102"].value or 0,
        "site_evening": ws["F102"].value or 0,
        "site_weekend": ws["G102"].value or 0,
    
        "outside_m25": ws["E106"].value or 0,
        "inside_m25": ws["E107"].value or 0,
        "mileage": ws["E108"].value or 0
    }


    return rates



def generate_pricing_excel(uploaded_file):

    uploaded_file.seek(0)
    file_bytes = uploaded_file.read()
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)

    ws = wb["PRICING SHEET"]

    # ✅ Header info
    ws["C5"] = sap_description
    ws["C8"] = consultant_name
    ws["J5"] = currency
    # ✅ SAP Sheet
    sap_ws = wb["SAP INFO FORM"]

    sap_ws["D7"] = customer_contact_name
    sap_ws["D8"] = contact_tel
    sap_ws["D9"] = contact_email

    sap_ws["D17"] = bfo_opp_no
    sap_ws["D18"] = material_code_count

    sap_ws["D19"] = material_code_1
    sap_ws["D20"] = material_code_2
    sap_ws["D21"] = material_code_3
    
    sap_ws["E19"] = material_price_1
    sap_ws["E20"] = material_price_2 if material_code_count >= 2 else ""
    sap_ws["E21"] = material_price_3 if material_code_count == 3 else ""


    # ==========================
    # ✅ BILLING MILESTONES
    # ==========================
    
    sap_ws["D37"] = billing_milestone_count
    
    for i in range(billing_milestone_count):
        
        value_cell = 38 + (i * 2)
        date_cell = 39 + (i * 2)
    
        sap_ws[f"D{value_cell}"] = billing_values[i]
        
        # Convert date to Excel format (string safe)
        sap_ws[f"D{date_cell}"] = billing_dates[i].strftime("%d/%m/%Y")




    # ==========================
    # ✅ LABOUR ITEMS (ROWS 15–24)
    # ==========================

    for i, lr in enumerate(st.session_state.labour_rows):

        if i >= 10:  # max rows (15–24)
            break

        row = 15 + i

        ws[f"C{row}"] = lr["description"]
        ws[f"D{row}"] = lr["office_day"]
        ws[f"E{row}"] = lr["site_day"]
        ws[f"F{row}"] = lr["office_evening"]
        ws[f"G{row}"] = lr["site_evening"]
        ws[f"H{row}"] = lr["office_weekend"]
        ws[f"I{row}"] = lr["site_weekend"]

    # ==========================
    # ✅ OTHER COSTS (ROWS 39–44)
    # ==========================

    for i, oc in enumerate(st.session_state.other_cost_rows):

        if i >= 6:  # max rows (39–44)
            break

        row = 39 + i

        ws[f"C{row}"] = oc["description"]
        ws[f"D{row}"] = oc["cost"]
        
        cell = ws[f"E{row}"]
        cell.value = None   # clear properly first
        cell.value = "GBP"

        ws[f"G{row}"] = oc["margin"] / 100

    # ==========================
    # ✅ EXPENSES
    # ==========================

    ws["D30"] = overnight_outside
    ws["D31"] = overnight_inside
    ws["D32"] = miles
    ws["D33"] = flights_cost
    ws["E33"] = "GBP"

    # ==========================
    # ✅ DISCOUNT
    # ==========================

    ws["D62"] = discount_pct / 100

    return wb

# ==========================
# SESSION STATE
# ==========================


if "works_list" not in st.session_state:
    st.session_state.works_list = []

if "other_cost_rows" not in st.session_state:
    st.session_state.other_cost_rows = []

if "labour_rows" not in st.session_state:
    st.session_state.labour_rows = []

if "price1" not in st.session_state:
    st.session_state.price1 = 0.0

if "bm_value_0" not in st.session_state:
    st.session_state.bm_value_0 = 0.0


if st.session_state.get("do_autofill", False):

    if "total_price" in st.session_state:
        st.session_state.price1 = float(st.session_state.total_price)
        st.session_state.bm_value_0 = float(st.session_state.total_price)

    st.session_state.do_autofill = False


if "currency" not in st.session_state:
    st.session_state.currency = "GBP"

currency = st.session_state.currency

if currency == "EUR":
    fx_rate = 1.16
    currency_symbol = "€"
else:
    fx_rate = 1
    currency_symbol = "£"

rates = {
    "office_day": 0,
    "site_day": 0,
    "office_evening": 0,
    "site_evening": 0,
    "office_weekend": 0,
    "site_weekend": 0
}

expenses_total = 0
discount_factor = 1


# ==========================
# PAGE SETUP
# ==========================
st.set_page_config(page_title="Energy Quote Tool", layout="centered")
st.title("Consultancy Quote Generator")

st.markdown("---")
# ==========================
# 📊 COST SHEET UPLOAD
# ==========================
st.subheader("📊 Cost Sheet")

uploaded_file = st.file_uploader(
    "Upload Pricing Template",
    type=["xlsx"]
)
st.markdown("---")

# ==========================
# DISCIPLINE SELECTION
# ==========================
st.subheader("Quotation Type")

discipline_options = ["Energy", "Power", "Microgrid", "Data Centre"]
discipline = st.selectbox("Select Discipline", discipline_options)


# ==========================
# QUOTE TYPE BY DISCIPLINE
# ==========================

quote_options_by_discipline = {

    "Energy": {
        "Energy Efficiency Audit": "templates/Energy Efficiency Audit Template.docx",
        "Metering Assessment": "templates/Metering Assessment Template.docx",
        "EE and Metering": "templates/EE Audit and Metering Template.docx",
        "ESOS P4": "templates/ESOS P4 Template.docx",
        "ESOS P4 and Transport": "templates/ESOS P4 and Transport Template.docx",
        "ESOS P4 Transport": "templates/ESOS P4 Transport Template.docx",
        "Energy Performance Certificate": "templates/Energy Performance Certificate Template.docx"
    },

    "Power": {},       # ✅ ready for future
    "Microgrid": {},   # ✅ ready for future
    "Data Centre": {}  # ✅ ready for future
}


quote_options = quote_options_by_discipline.get(discipline, {})

if quote_options:
    quote_type = st.selectbox("Select Quote Type", list(quote_options.keys()))
    st.session_state.selected_quote_type = quote_type
else:
    st.warning("No quote types available for this discipline yet.")
    quote_type = None

st.markdown("---")

# ==========================
# INPUT FIELDS
# ==========================

with st.expander("📄 Template Info", expanded=False):

    col1, col2 = st.columns(2)
    
    with col1:
        customer_name = st.text_input("Customer Name")
        number_of_sites = st.text_input("Number of Sites")
    
        site_name = st.text_input(
            "Site Name",
            placeholder="e.g. Coventry, London and Warrington"
        )
        st.caption("If multiple sites, use this format: Coventry, London and Warrington")
    
    with col2:
        project_name = st.text_input("Project Name")
    
        consultant_name = st.text_input(
            "Consultant Name",
        )
        number_of_consultants = st.number_input(
            "Number of Consultants",
            min_value=1,
            max_value=50,
            value=1
        )

# ==========================
# 🏢 CUSTOMER DETAILS
# ==========================
st.markdown("---")
with st.expander("🏢 Customer Details", expanded=False):

    col1, col2 = st.columns(2)
    
    with col1:
        address_line_1 = st.text_input("Address Line 1")
        address_line_2 = st.text_input("Address Line 2")
    
    with col2:
        city = st.text_input("City")
        postcode = st.text_input("Postcode")
    
    # Contact name (optional)
    contact_name = st.text_input("Contact Name (leave blank for default)")

# ==========================
# 💰 LIVE COST CALCULATOR
# ==========================
if uploaded_file is not None:
    with st.expander("💰 Live Cost Calculator", expanded=True):
        rates = extract_rates(uploaded_file)
        st.write("DEBUG RATES:", rates)
    
        st.markdown("---")
        st.subheader("💰 Live Cost Calculator")
    
        currency = st.selectbox(
            "Currency",
            ["GBP", "EUR"],
            key="currency"
        )
        
        currency = st.session_state.currency
        
        if currency == "EUR":
            fx_rate = 1.16
            currency_symbol = "€"
        else:
            fx_rate = 1
            currency_symbol = "£"
    
        # Inputs
        st.markdown("### Labour Hours (Detailed)")
    
    # ✅ Add row button
        if st.button("➕ Add Labour Row"):
            st.session_state.labour_rows.append({
                "description": "",
                "office_day": 0.0,
                "site_day": 0.0,
                "office_evening": 0.0,
                "site_evening": 0.0,
                "office_weekend": 0.0,
                "site_weekend": 0.0
            })
        
        # ✅ Display rows
        total_office_day = 0
        total_site_day = 0
        total_office_evening = 0
        total_site_evening = 0
        total_office_weekend = 0
        total_site_weekend = 0
        
        for i, row in enumerate(st.session_state.labour_rows):
        
            st.markdown(f"#### Work Item {i+1}")
        
            row["description"] = st.text_input(
                "Description",
                value=row["description"],
                key=f"payment_desc_{i}"
            )
        
            col1, col2, col3 = st.columns(3)
        
            
            with col1:
                row["office_day"] = st.number_input(
                    "Office Hours (Mon-Fri 0800 to 1700)",
                    0.0,
                    key=f"od_{i}"
                )
                row["site_day"] = st.number_input(
                    "On Site Hours (Mon-Fri 0800 to 1700)",
                    0.0,
                    key=f"sd_{i}"
                )
            
            with col2:
                row["office_evening"] = st.number_input(
                    "Office Hours (Mon-Fri 1700 to 2400)",
                    0.0,
                    key=f"oe_{i}"
                )
                row["site_evening"] = st.number_input(
                    "On Site Hours (Mon-Fri 1700 to 2400)",
                    0.0,
                    key=f"se_{i}"
                )
            
            with col3:
                row["office_weekend"] = st.number_input(
                    "Office Hours (Sat&Sun 0800 to 2400)",
                    0.0,
                    key=f"ow_{i}"
                )
                row["site_weekend"] = st.number_input(
                    "On Site Hours (Sat & Sun 0800 to 2400)",
                    0.0,
                    key=f"sw_{i}"
                )
    
        
            # ✅ Accumulate totals
            total_office_day += row["office_day"]
            total_site_day += row["site_day"]
            total_office_evening += row["office_evening"]
            total_site_evening += row["site_evening"]
            total_office_weekend += row["office_weekend"]
            total_site_weekend += row["site_weekend"]
        
            st.markdown("---")
       
        st.markdown("### Other Costs (Detailed)")
    
        # ✅ Add row button
        if st.button("➕ Add Cost Row"):
            st.session_state.other_cost_rows.append({
                "description": "",
                "cost": 0.0,
                "margin": 0.0,
                "selling": 0.0
            })
        
        
        total_other_cost = 0
        total_other_selling = 0
        
        # ✅ Display rows
        for i, row in enumerate(st.session_state.other_cost_rows):
        
            st.markdown(f"#### Cost Item {i+1}")
        
            row["description"] = st.text_input(
                "Description",
                value=row["description"],
                key=f"other_desc_{i}"
            )
        
            col1, col2 = st.columns(2)
        
            with col1:
                row["cost"] = st.number_input(
                    "Cost (£)",
                    0.0,
                    key=f"other_cost_{i}"
                )
        
            with col2:
                row["margin"] = st.number_input(
                    "Margin (%)",
                    0.0,
                    100.0,
                    key=f"other_margin_{i}"
                )
        
            # ✅ Selling calculation
            if row["margin"] < 100:
                row["selling"] = row["cost"] / (1 - row["margin"] / 100)
            else:
                row["selling"] = 0
        
            st.write(f"Selling Price: £{row['selling']:,.2f}")
        
            total_other_cost += row["cost"]
            total_other_selling += row["selling"]
        
            st.markdown("---")
        
        # ✅ Totals (like Excel bottom row)
    
        peer_review_hours = 0.1 * total_office_day
        
        st.write("### Totals")
        st.write(f"Office Day: {total_office_day}")
        st.write(f"Peer Review (Office): {peer_review_hours}")
        st.write(f"Site Day: {total_site_day}")
        
        st.write(f"Office Evening: {total_office_evening}")
        st.write(f"Site Evening: {total_site_evening}")
        
        st.write(f"Office Weekend: {total_office_weekend}")
        st.write(f"Site Weekend: {total_site_weekend}")
        
        # ✅ SELLING LABOUR (NEW)
        labour_total = (
            (total_office_day * rates["office_day"]) +
            (total_site_day * rates["site_day"]) +
        
            (total_office_evening * rates["office_evening"]) +
            (total_site_evening * rates["site_evening"]) +
        
            (total_office_weekend * rates["office_weekend"]) +
            (total_site_weekend * rates["site_weekend"])
        )
        
        # ✅ Peer review
        peer_review = 0.1 * total_office_day * rates["office_day"]
        peer_review_factor = 0.1
        
        labour_total += peer_review
    
        st.markdown("### Expenses & Costs")
    
        col1, col2 = st.columns(2)
        
        with col1:
            overnight_outside = st.number_input("Overnight Stays (Outside M25)", 0)
            overnight_inside = st.number_input("Overnight Stays (Inside M25)", 0)
            miles = st.number_input("Mileage (miles)", 0)
        
        with col2:
            flights_cost = st.number_input("Flights / Rail (£)", 0.0)
            discount_pct = st.number_input("Discount (%)", 0.0, 100.0, 0.0)
            discount_factor = 1 - (discount_pct / 100)
    
        # ✅ EXPENSES (SELLING)
        expenses_total = (
            overnight_outside * rates.get("outside_m25", 0)
            + overnight_inside * rates.get("inside_m25", 0)
            + miles * rates.get("mileage", 0)
            + flights_cost * 1.15
        )
    
            
        # ==========================
        # Other Costs
        # ==========================
            
        
        other_cost = total_other_cost
        other_cost_selling = total_other_selling
    
            
        # ==========================
        # Final Total
        # ==========================
    
            
        # ✅ COST CALCULATION (from Excel logic)
        
        labour_cost = (
            (total_office_day * rates["office_cost"]) +
            (total_site_day * rates["site_cost"]) +
        
            (total_office_evening * rates["office_cost"]) +
            (total_site_evening * rates["site_cost"]) +
        
            (total_office_weekend * rates["office_cost"]) +
            (total_site_weekend * rates["site_cost"])
        )
        
        peer_review_cost = 0.1 * total_office_day * rates["office_cost"]
        
        labour_cost += peer_review_cost
    
    
        # ✅ Expense COST (not selling)
        expenses_cost = (
            overnight_outside * (rates.get("outside_m25", 0) / 1.15 if rates.get("outside_m25", 0) else 0)
            + overnight_inside * (rates.get("inside_m25", 0) / 1.15 if rates.get("inside_m25", 0) else 0)
            + miles * (rates.get("mileage", 0) / 1.675 if rates.get("mileage", 0) else 0)
            + flights_cost
        )
    
        total_cost = labour_cost + expenses_cost + other_cost
    
            
        subtotal = labour_total + expenses_total + other_cost_selling
            
        discount_value = subtotal * (discount_pct / 100)
            
        total_price = subtotal - discount_value
        st.session_state.total_price = total_price
    
            
        labour_total_fx = labour_total * fx_rate
        expenses_total_fx = expenses_total * fx_rate
        other_cost_selling_fx = other_cost_selling * fx_rate
        
        total_cost_fx = total_cost * fx_rate
        subtotal_fx = subtotal * fx_rate
        total_price_fx = total_price * fx_rate
        labour_cost_fx = labour_cost * fx_rate
        expenses_cost_fx = expenses_cost * fx_rate
        other_cost_fx = other_cost * fx_rate
    
        labour_selling_discounted = labour_total * discount_factor
        expenses_selling_discounted = expenses_total * discount_factor
        other_selling_discounted = other_cost_selling * discount_factor
    
        labour_selling_discounted_fx = labour_selling_discounted * fx_rate
        expenses_selling_discounted_fx = expenses_selling_discounted * fx_rate
        other_selling_discounted_fx = other_selling_discounted * fx_rate
    
        if subtotal > 0:
            margin_pct = (subtotal - total_cost) / subtotal * 100
        else:
            margin_pct = 0
    
        if total_price > 0:
            actual_margin_pct = (total_price - total_cost) / total_price * 100
        else:
            actual_margin_pct = 0
    
        
        st.markdown("### Breakdown")
    
        st.markdown("#### Labour")
        st.write(f"Selling: {currency_symbol}{labour_total_fx:,.2f}")
        st.write(f"Cost: {currency_symbol}{labour_cost_fx:,.2f}")
        st.write(f"Selling After Discount: {currency_symbol}{labour_selling_discounted_fx:,.2f}")
        
        st.markdown("---")
        
        st.markdown("#### Expenses")
        st.write(f"Selling: {currency_symbol}{expenses_total_fx:,.2f}")
        st.write(f"Cost: {currency_symbol}{expenses_cost_fx:,.2f}")
        st.write(f"Selling After Discount: {currency_symbol}{expenses_selling_discounted_fx:,.2f}")
        
        st.markdown("---")
        
        st.markdown("#### Other Costs")
        st.write(f"Selling: {currency_symbol}{other_cost_selling_fx:,.2f}")
        st.write(f"Cost: {currency_symbol}{other_cost_fx:,.2f}")
        st.write(f"Selling After Discount: {currency_symbol}{other_selling_discounted_fx:,.2f}")
        
        st.markdown("---")
          
    
        st.write(f"Total Cost: {currency_symbol}{total_cost_fx:,.2f}")
        st.write(f"Selling Price: {currency_symbol}{subtotal_fx:,.2f}")
    
        st.write(f"Margin (%): {margin_pct:.2f}%")
        
        st.write(f"Discount (%): {discount_pct:.2f}%")
        st.write(f"Actual Selling Price: {currency_symbol}{total_price_fx:,.2f}")
        st.write(f"Actual Margin (%): {actual_margin_pct:.2f}%")
        
        st.markdown("---")
        st.metric("Total Price", f"{currency_symbol}{total_price_fx:,.2f}")
    
        
        if st.button("Auto Fill Pricing Fields"):
            st.session_state["do_autofill"] = True

# ==========================
# WORKS INPUT
# ==========================
with st.expander("🛠️ Works & Pricing", expanded=True):
    combined_items = []
    
    # Labour
    for lr in st.session_state.labour_rows:
        if lr["description"]:
            combined_items.append({
                "description": lr["description"],
                "type": "labour",
                "data": lr
            })
    
    # Other Costs
    for oc in st.session_state.other_cost_rows:
        if oc["description"]:
            combined_items.append({
                "description": oc["description"],
                "type": "other",
                "data": oc
            })
    
    # Expenses
    if 'expenses_total' in locals() and expenses_total > 0:
        combined_items.append({
            "description": "Expenses",
            "type": "expenses"
        })
    
    
    # ✅ AUTO-POPULATE ONLY IF EMPTY OR SIZE CHANGED
    if len(st.session_state.works_list) != len(combined_items):
        
        st.session_state.works_list = []
    
        for item in combined_items:
            st.session_state.works_list.append({
                "description": item["description"],
                "mode": "Auto",
                "include_labour": True,
                "include_other": True,
                "include_expenses": True,
                "manual_price": 0.0
            })
    
    total_works_price = 0
    
    for i, work in enumerate(st.session_state.works_list):
    
        
        col1, col2 = st.columns([6, 1])
        
        with col1:
            st.markdown(f"#### Work Item {i+1}")
        
        with col2:
            if st.button("❌", key=f"delete_{i}"):
                st.session_state.works_list.pop(i)
                st.rerun()
    
            if uploaded_file is None:
                work["mode"] = "Manual"
                st.selectbox(
                    "Pricing Mode",
                    ["Manual"],
                    key=f"mode_{i}"
                )
            else:
                work["mode"] = st.selectbox(
                    "Pricing Mode",
                    ["Auto", "Manual"],
                    key=f"mode_{i}"
                )
    
        # ==========================
        # ✅ AUTO MODE
        # ==========================
        if work["mode"] == "Auto" and uploaded_file is not None:
    
            price = 0
    
            combined_items = []
    
            # ✅ Labour rows
            for lr in st.session_state.labour_rows:
                if lr["description"]:
                    combined_items.append({
                        "description": lr["description"],
                        "type": "labour",
                        "data": lr
                    })
    
            # ✅ Other costs
            for oc in st.session_state.other_cost_rows:
                if oc["description"]:
                    combined_items.append({
                        "description": oc["description"],
                        "type": "other",
                        "data": oc
                    })
    
            # ✅ Expenses
            if expenses_total > 0:
                combined_items.append({
                    "description": "Expenses",
                    "type": "expenses"
                })
    
            # ✅ Assign item based on position
            if i < len(combined_items):
                item = combined_items[i]
                work["description"] = item["description"]
            
                if item["type"] == "labour":
                    lr = item["data"]
            
                    price = (
                        # ✅ Normal labour
                        lr["office_day"] * rates["office_day"] +
                        lr["site_day"] * rates["site_day"] +
                        lr["office_evening"] * rates["office_evening"] +
                        lr["site_evening"] * rates["site_evening"] +
                        lr["office_weekend"] * rates["office_weekend"] +
                        lr["site_weekend"] * rates["site_weekend"]
            
                        # ✅ Peer review (correctly included ✅)
                        + (lr["office_day"] * rates["office_day"] * 0.1)
            
                    ) * discount_factor
    
                elif item["type"] == "other":
                    price = item["data"]["selling"] * discount_factor
    
                elif item["type"] == "expenses":
                    price = expenses_total * discount_factor
    
            else:
                work["description"] = ""
                price = 0
    
            # ✅ Locked description display
            st.text_input(
                "Description",
                value=work["description"],
                key=f"auto_desc_locked_{i}",
                disabled=True
            )
    
        # ==========================
        # ✅ MANUAL MODE
        # ==========================
        else:
            work["description"] = st.text_input(
                "Description",
                value=work["description"],
                key=f"work_desc_{i}"
            )
    
            work["manual_price"] = st.number_input(
                "Manual Price (£)",
                0.0,
                key=f"manual_{i}"
            )
    
            price = work["manual_price"]
    
        # ✅ FINAL PRICE DISPLAY (ONLY ONCE)
        work["price"] = price
        
        price_fx = price * fx_rate
        st.write(f"Price: {currency_symbol}{price_fx:,.2f}")
    
        total_works_price += price
    
        st.markdown("---")
    
    
    st.write("### Works Total")
    total_works_price_fx = total_works_price * fx_rate
    st.write(f"Total Works Price: {currency_symbol}{total_works_price_fx:,.2f}")


# ==========================
# 💰 PAYMENT TERMS
# ==========================
st.markdown("---")

with st.expander("💰 Payment Terms", expanded=False):

    if "payment_terms" not in st.session_state:
        st.session_state.payment_terms = [
            {"percent": 100, "description": "upon submittal of the report"}
        ]
    
    # Display inputs
    for i, term in enumerate(st.session_state.payment_terms):
        col1, col2 = st.columns([1, 3])
    
        with col1:
            st.session_state.payment_terms[i]["percent"] = st.number_input(
                f"% {i+1}",
                min_value=0,
                max_value=100,
                value=term["percent"],
                key=f"percent_{i}"
            )
    
        with col2:
            st.session_state.payment_terms[i]["description"] = st.text_input(
                f"Condition {i+1}",
                value=term["description"],
                key=f"desc_{i}"
            )
    
    # Add new row
    if st.button("➕ Add Payment Split"):
        st.session_state.payment_terms.append({"percent": 0, "description": ""})
    
    # Calculate total
    total_percent = sum(term["percent"] for term in st.session_state.payment_terms)
    
    if total_percent != 100:
        st.warning(f"⚠️ Total must equal 100% (Currently {total_percent}%)")
    else:
        st.success("✅ Payment terms total = 100%")

# ==========================
# 📋 PRICING SHEET INFO
# ==========================
if uploaded_file is not None:

    st.markdown("---")
    with st.expander("📋 Pricing Sheet Info", expanded=False):

        col1, col2 = st.columns(2)
    
        with col1:
            sap_description = st.text_input("SAP Description Name")  
            bfo_opp_no = st.text_input("bFO Opportunity No")
    
        # ✅ Keep these inside too (important)
        customer_contact_name = st.text_input("Customer Contact Name")
        contact_tel = st.text_input("Contact Tel No")
        contact_email = st.text_input("Contact Email")
    
        with col2:
    
            material_code_count = st.number_input(
                "No. of Material Codes",
                min_value=1,
                max_value=3,
                value=1
            )
    
            material_code_1 = st.selectbox(
                "Material Code No.1",
                material_code_options,
                key="mc1"
            )
    
            material_price_1 = st.number_input(
                "Price for Material Code 1 (£)",
                min_value=0.0,
                key="price1"
            )
    
            # ✅ Material Code 2
            material_code_2 = ""
            material_price_2 = 0.0
    
            if material_code_count >= 2:
                material_code_2 = st.selectbox(
                    "Material Code No.2",
                    material_code_options,
                    key="mc2"
                )
    
                material_price_2 = st.number_input(
                    "Price for Material Code 2 (£)",
                    min_value=0.0,
                    key="price2"
                )
    
            # ✅ Material Code 3
            material_code_3 = ""
            material_price_3 = 0.0
    
            if material_code_count == 3:
                material_code_3 = st.selectbox(
                    "Material Code No.3",
                    material_code_options,
                    key="mc3"
                )
    
                material_price_3 = st.number_input(
                    "Price for Material Code 3 (£)",
                    min_value=0.0,
                    key="price3"
                )


# ==========================
# 💰 BILLING MILESTONES (ONLY IF FILE UPLOADED)
# ==========================
if uploaded_file is not None:

    with st.expander("📆 Billing Milestones", expanded=False):

        billing_milestone_count = st.selectbox(
            "No. of Billing Milestones",
            [1, 2, 3, 4, 5]
        )
    
        billing_values = []
        billing_dates = []
    
        for i in range(billing_milestone_count):
            st.markdown(f"#### Billing Milestone {i+1}")
    
            value = st.number_input(
                f"Milestone {i+1} Value (£)",
                min_value=0.0,
                key=f"bm_value_{i}"
            )
    
            date = st.date_input(
                f"Milestone {i+1} Planned Billing Date",
                key=f"bm_date_{i}"
            )
    
            billing_values.append(value)
            billing_dates.append(date)

# ==========================
# 📄 FILE NAMING SECTION
# ==========================
st.markdown("---")
with st.expander("📄 File Naming", expanded=False):

    document_type_options = {
        "Cost Sheet": "CST",
        "Quote": "QTE",
        "Calculation": "CLC",
        "Data Collection Form": "DCF",
        "Document": "DOC",
        "Drawing": "DRG",
        "Functional Design Specification": "FDS",
        "Risk Assessment / Method Statement": "RMS",
        "Report": "RPT",
        "Schedule": "SCH",
        "Specification": "SPC"
    }
    
    subject_options = {
        "Audit (EcoConsult Audit for Power)": "ADT",
        "Block": "BLK",
        "Cabling": "CBL",
        "Design": "DES",
        "Data Centre Audit": "DTC",
        "Equipment": "EQP",
        "Earthing": "ETH",
        "Factory Acceptance Test": "FAT",
        "Feasibility": "FSY",
        "General Arrangement or Layout": "GAR",
        "Energy Audit": "NRG",
        "Microgrid Feasibility": "MGF",
        "Microgrid Design": "MGD",
        "Metering Survey": "MTR",
        "Protection": "PRT",
        "Power Quality": "PQT",
        "Pressure Rise Study": "PRS",
        "Power System Study": "PSS"
    }
    
    col1, col2 = st.columns(2)
    
    with col1:
        
            # ✅ Always sync Project Number to Project Name (LIVE)
        st.session_state.project_number = project_name
        
        # ✅ Live sync Project Name → Project Number
        st.session_state.project_number = project_name
        
        project_number = st.text_input(
            "Project Number",
            key="project_number",
            value=st.session_state.project_number
        )
    
    
        document_type_label = st.selectbox("Document Type", list(document_type_options.keys()))
        document_type = document_type_options[document_type_label]
    
        subject_label = st.selectbox("Subject", list(subject_options.keys()))
        subject = subject_options[subject_label]
    
    with col2:
        unique_id = st.selectbox("Unique Identifier", [f"{i:02d}" for i in range(1, 21)])
    
        revision_code_label = st.selectbox(
            "Revision Code",
            ["Contractual (External)", "Preliminary (Internal)"]
        )
        revision_code = "C" if "Contractual" in revision_code_label else "P"
    
        revision_number = st.selectbox("Revision Number", [f"{i:02d}" for i in range(1, 21)])

# ==========================
# TRANSPORT SECTION
# ==========================
st.markdown("---")

if "Transport" in quote_type:
    with st.expander("🚚 Transport Details", expanded=False):

        col1, col2 = st.columns(2)
    
        with col1:
            number_of_transport = st.text_input("Number of Transports")
    
        with col2:
            transport_type = st.text_input(
                "Transport Type",
                placeholder="e.g. HGV and Grey Fleet"
            )
            st.caption("If multiple types, use: HGV and Grey Fleet")
    
else:
    number_of_transport = ""
    transport_type = ""

st.markdown("---")

# ==========================
# ✅ SAFE PLACEHOLDER FUNCTION
# ==========================
   
def replace_placeholders(doc, data):

    def process_paragraph(paragraph):

        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"

            # ✅ Check full text
            full_text = "".join(run.text for run in paragraph.runs)

            if placeholder in full_text:

                # ✅ Replace across runs carefully
                remaining_text = full_text.replace(placeholder, str(value))

                # ✅ Write back character by character across runs
                i = 0
                for run in paragraph.runs:
                    run_len = len(run.text)

                    run.text = remaining_text[i:i+run_len]

                    i += run_len

                # ✅ If any text left, append to last run
                if i < len(remaining_text):
                    paragraph.runs[-1].text += remaining_text[i:]

    # ✅ Process paragraphs
    for paragraph in doc.paragraphs:
        process_paragraph(paragraph)

    # ✅ Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    process_paragraph(paragraph)

    return doc


def fill_works_table(doc, works_list, fx_rate, currency_symbol):

    target_table = None

    # ✅ Find the table with "Work Description" header
    for table in doc.tables:
        first_row = table.rows[0].cells

        if "Description" in first_row[1].text:
            target_table = table
            break

    if target_table is None:
        print("❌ Pricing table not found")
        return doc

    start_row = 1

    # ✅ Fill rows
    for i, work in enumerate(works_list):

        if start_row + i >= len(target_table.rows) - 1:
            row_cells = target_table.add_row().cells
        else:
            row_cells = target_table.rows[start_row + i].cells

        row_cells[0].text = str(i + 1)
        row_cells[1].text = work["description"]
        price_fx = work["price"] * fx_rate
        row_cells[2].text = f"{currency_symbol}{price_fx:,.2f}"

    # ✅ Total row
    total = sum(work["price"] for work in works_list)
    total_fx = total * fx_rate
    target_table.rows[-1].cells[2].text = f"{currency_symbol}{total_fx:,.2f}"

    return doc



def insert_payment_terms(doc, payment_terms):

    payment_lines = []

    for term in payment_terms:
        if term["percent"] > 0 and term["description"]:
            payment_lines.append(f" {term['percent']}% {term['description']}")

    payment_text = "\n".join(payment_lines)

    # ✅ 1. Normal paragraphs
    for paragraph in doc.paragraphs:
        if "{{PaymentTerms}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{PaymentTerms}}", payment_text)

    # ✅ 2. TABLE CELLS (THIS IS THE MISSING PIECE)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if "{{PaymentTerms}}" in paragraph.text:
                        paragraph.text = paragraph.text.replace("{{PaymentTerms}}", payment_text)

    return doc



# ==========================
# GENERATE DOCUMENT
# ==========================
if st.button("📄 Generate Word Document"):
    
    
    if uploaded_file is not None and "office_hours" in locals():
        
        uploaded_file.seek(0)
        file_bytes = uploaded_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        ws = wb["PRICING SHEET"]
    
        # ==========================
        # ✅ PRICING SHEET HEADER INFO
        # ==========================
        
        ws["C5"] = sap_description
        ws["C8"] = consultant_name
        ws["J5"] = currency
        ws["B10"] = office_hours
        ws["C10"] = site_hours
        # ==========================
        # ✅ SAP INFO FORM SHEET
        # ==========================
        
        sap_ws = wb["SAP INFO FORM"]
        
        sap_ws["D7"] = customer_contact_name
        sap_ws["D8"] = contact_tel
        sap_ws["D9"] = contact_email
        
        sap_ws["D17"] = bfo_opp_no
        sap_ws["D18"] = material_code_count
        
        sap_ws["D19"] = material_code_1
        sap_ws["D20"] = material_code_2
        sap_ws["D21"] = material_code_3


    if not customer_name or not project_name:
        st.error("Customer Name and Project Name are required.")

    elif not st.session_state.works_list:
        st.error("Please add at least one work item before generating the document.")

    elif not st.session_state.works_list:
        st.error("Please add at least one work item before generating the document.")
    
    elif any(
        not work.get("description") or work.get("price", 0) == 0
        for work in st.session_state.works_list
    ):
        st.error("❌ All work items must have a description and non-zero price before generating the document.")

    else:
        template_path = quote_options[quote_type]
        doc = Document(template_path)

        # ✅ Fix contact name logic
        if not contact_name.strip():
            contact_name_final = "whom it may concern"
        else:
            contact_name_final = contact_name

        # ✅ Build address properly (FIXED INDENT)
        address_lines = []

        if address_line_1:
            address_lines.append(address_line_1)

        if address_line_2:
            address_lines.append(address_line_2)

        if city:
            address_lines.append(city)

        if postcode:
            address_lines.append(postcode)

        full_address = "\n".join(address_lines)

        # ✅ Build data dictionary
        data = {
            "CustomerName": customer_name,
            "NumberOfSites": number_of_sites,
            "SiteName": site_name,
            "NumberOfTransport": number_of_transport,
            "TransportType": transport_type,
            "ProjectName": project_name,
            "NumberOfConsultants": number_of_consultants,

            "FullAddress": full_address,
            "ContactName": contact_name_final,

            "TodaysDate": datetime.now().strftime("%d %B %Y")
        }

        # ✅ Apply transformations
        doc = insert_payment_terms(doc, st.session_state.payment_terms)
        doc = replace_placeholders(doc, data)

        if st.session_state.works_list:
            doc = fill_works_table(doc, st.session_state.works_list, fx_rate, currency_symbol)

        # ✅ Filename
        file_name = f"{project_number}-{document_type}-{subject}-{unique_id}-{revision_code}{revision_number}"

        output = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(output.name)

        st.success("✅ Document generated successfully!")

        with open(output.name, "rb") as f:
            st.download_button(
                "⬇ Download Quote",
                f,
                file_name=f"{file_name}.docx"
            )
        
        # ✅ Generate Excel
        excel_wb = generate_pricing_excel(uploaded_file)
        
        excel_output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        excel_wb.save(excel_output.name)
        
        with open(excel_output.name, "rb") as f:
            st.download_button(
                "⬇ Download Pricing Sheet",
                f,
                file_name=f"{project_number}-PricingSheet.xlsx"
            )

