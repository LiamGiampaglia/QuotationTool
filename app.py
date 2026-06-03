from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
import streamlit as st
from docx import Document
import tempfile
from datetime import datetime
import openpyxl

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


# ==========================
# SESSION STATE
# ==========================


if "works_list" not in st.session_state:
    st.session_state.works_list = []

if "other_cost_rows" not in st.session_state:
    st.session_state.other_cost_rows = []

if "labour_rows" not in st.session_state:
    st.session_state.labour_rows = []


# ==========================
# PAGE SETUP
# ==========================
st.set_page_config(page_title="Energy Quote Tool", layout="centered")
st.title("Consultancy Quote Generator")


# ==========================
# 📊 COST SHEET UPLOAD
# ==========================
st.markdown("---")
st.subheader("📊 Cost Sheet")

uploaded_file = st.file_uploader(
    "Upload Pricing Template",
    type=["xlsx"]
)


# ==========================
# DISCIPLINE SELECTION
# ==========================

discipline_options = ["Energy", "Power", "Microgrid", "Data Centre"]
discipline = st.selectbox("Select Discipline", discipline_options)

st.markdown("---")


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
        "ESOS P4 Transport": "templates/ESOS P4 Transport Template.docx"
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
st.subheader("🏢 Customer Details")

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
# 📄 FILE NAMING SECTION
# ==========================
st.markdown("---")
st.subheader("📄 File Naming")

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
    project_number = st.text_input("Project Number")

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
    st.subheader("🚚 Transport Details")

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
# 💰 LIVE COST CALCULATOR
# ==========================

if uploaded_file is not None:

    rates = extract_rates(uploaded_file)
    st.write("DEBUG RATES:", rates)

    st.markdown("---")
    st.subheader("💰 Live Cost Calculator")

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
            row["office_day"] = st.number_input("Office Day", 0.0, key=f"od_{i}")
            row["site_day"] = st.number_input("Site Day", 0.0, key=f"sd_{i}")
    
        with col2:
            row["office_evening"] = st.number_input("Office Evening", 0.0, key=f"oe_{i}")
            row["site_evening"] = st.number_input("Site Evening", 0.0, key=f"se_{i}")
    
        with col3:
            row["office_weekend"] = st.number_input("Office Weekend", 0.0, key=f"ow_{i}")
            row["site_weekend"] = st.number_input("Site Weekend", 0.0, key=f"sw_{i}")
    
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
    
    labour_total += peer_review

    st.markdown("### Expenses & Costs")

    col1, col2 = st.columns(2)
    
    with col1:
        overnight_outside = st.number_input("Overnight Stays (Outside M25)", 0)
        overnight_inside = st.number_input("Overnight Stays (Inside M25)", 0)
        miles = st.number_input("Mileage (miles)", 0)
    
    with col2:
        flights_cost = st.number_input("Flights / Rail (£)", 0.0)
        other_cost = st.number_input("Other Costs (£)", 0.0)
        discount_pct = st.number_input("Discount (%)", 0.0, 100.0, 0.0)

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


    if subtotal > 0:
        margin_pct = (subtotal - total_cost) / subtotal * 100
    else:
        margin_pct = 0

    if total_price > 0:
        actual_margin_pct = (total_price - total_cost) / total_price * 100
    else:
        actual_margin_pct = 0


    # ✅ Display (NOW OUTSIDE the if/else)
    st.markdown("### Breakdown")
    
    st.write(f"Labour: £{labour_total:,.2f}")
    st.write(f"Expenses: £{expenses_total:,.2f}")
    st.write(f"Other Costs: £{other_cost_selling:,.2f}")
    
    st.markdown("---")
    
    st.write(f"Total Cost: £{total_cost:,.2f}")
    st.write(f"Selling Price: £{subtotal:,.2f}")
    st.write(f"Margin (%): {margin_pct:.2f}%")
    
    st.write(f"Discount (%): {discount_pct:.2f}%")
    st.write(f"Actual Selling Price: £{total_price:,.2f}")
    st.write(f"Actual Margin (%): {actual_margin_pct:.2f}%")
    
    st.markdown("---")
    st.metric("Total Price", f"£{total_price:,.2f}")


# ==========================
# WORKS INPUT
# ==========================
st.subheader("🛠️ Works & Pricing")

if st.button("➕ Add Work"):
    st.session_state.works_list.append({
        "description": "",
        "mode": "Auto",
        "include_labour": True,
        "include_other": True,
        "include_expenses": True,
        "manual_price": 0.0
    })

total_works_price = 0

for i, work in enumerate(st.session_state.works_list):

    st.markdown(f"#### Work Item {i+1}")

    work["description"] = st.text_input(
        "Description",
        value=work["description"],
        key=f"work_desc_{i}"
    )

    work["mode"] = st.selectbox(
        "Pricing Mode",
        ["Auto", "Manual"],
        key=f"mode_{i}"
    )

    if work["mode"] == "Auto":
    
        # ✅ Select what this work item represents
        category = st.selectbox(
            "Auto Category",
            ["Energy Consultancy", "Other Costs", "Expenses"],
            key=f"cat_{i}"
        )
    
        # ✅ Set description + price automatically
        if category == "Energy Consultancy":
            work["description"] = "Energy Consultancy"
            price = labour_total
    
        elif category == "Other Costs":
            work["description"] = "Other Costs"
            price = other_cost_selling
    
        elif category == "Expenses":
            work["description"] = "Expenses"
            price = expenses_total
    
        # ✅ Show locked description
        st.text_input(
            "Description",
            value=work["description"],
            key=f"auto_desc_locked_{i}",
            disabled=True
        )



        # ✅ Reset
        description_lines = []
        price = 0
        
        # ✅ Labour → "Energy Consultancy"
        if work["include_labour"] and labour_total > 0:
            description_lines.append(f"Energy Consultancy – £{labour_total:,.2f}")
            price += labour_total
        
        # ✅ Other Costs
        if work["include_other"] and other_cost_selling > 0:
            description_lines.append(f"Other Costs – £{other_cost_selling:,.2f}")
            price += other_cost_selling
        
        # ✅ Expenses
        if work["include_expenses"] and expenses_total > 0:
            description_lines.append(f"Expenses – £{expenses_total:,.2f}")
            price += expenses_total


        # ✅ Build description text
        auto_description = "\n".join(description_lines)

        st.text_area(
            "Auto Description",
            value=auto_description,
            key=f"auto_desc_{i}",
            height=150
        )

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

    st.write(f"Price: £{price:,.2f}")

    total_works_price += price

    st.markdown("---")

st.write("### Works Total")
st.write(f"Total Works Price: £{total_works_price:,.2f}")


# ==========================
# 💰 PAYMENT TERMS
# ==========================
st.markdown("---")
st.subheader("💰 Payment Terms")

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


def fill_works_table(doc, works_list):

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
        row_cells[2].text = f"£{work['price']:,.2f}"

    # ✅ Total row
    total = sum(work["price"] for work in works_list)
    target_table.rows[-1].cells[2].text = f"£{total:,.2f}"

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
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb["PRICING SHEET"]
    
        ws["B10"] = office_hours
        ws["C10"] = site_hours

    if not customer_name or not project_name:
        st.error("Customer Name and Project Name are required.")

    elif not st.session_state.works_list:
        st.error("Please add at least one work item before generating the document.")

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
            doc = fill_works_table(doc, st.session_state.works_list)

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
