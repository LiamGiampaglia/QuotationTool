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

    rates = {}

    for row in ws.iter_rows(values_only=True):

        for i, cell in enumerate(row):

            if cell == "Office (GBP)" and i+1 < len(row):
                rates["office_cost"] = row[i+1]

            if cell == "Site (GBP)" and i+1 < len(row):
                rates["site_cost"] = row[i+1]

            if cell == "Mileage cost (GBP)" and i+1 < len(row):
                rates["mileage"] = row[i+1]

            if cell == "Outside M25 (GBP)" and i+1 < len(row):
                rates["outside_m25"] = row[i+1]

            if cell == "Inside M25 (GBP)" and i+1 < len(row):
                rates["inside_m25"] = row[i+1]

            if cell == "Margin In Office" and i+1 < len(row):
                rates["office_margin"] = row[i+1]

            if cell == "Margin On-Site" and i+1 < len(row):
                rates["site_margin"] = row[i+1]

    
            if cell == "Office (GBP)" and i+2 < len(row):
                rates["office_selling"] = row[i+2]
            
            if cell == "Site (GBP)" and i+2 < len(row):
                rates["site_selling"] = row[i+2]


    return rates

# ==========================
# SESSION STATE
# ==========================
if "works_list" not in st.session_state:
    st.session_state.works_list = []

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
    st.markdown("### Labour Hours (Match Excel)")

    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Mon–Fri 08:00–17:00**")
        office_day = st.number_input("Office (Day)", 0.0)
        site_day = st.number_input("Site (Day)", 0.0)
    
        st.write("**Mon–Fri 17:00–24:00**")
        office_evening = st.number_input("Office (Evening)", 0.0)
        site_evening = st.number_input("Site (Evening)", 0.0)
    
    with col2:
        st.write("**Weekend**")
        office_weekend = st.number_input("Office (Weekend)", 0.0)
        site_weekend = st.number_input("Site (Weekend)", 0.0)


    overnight_outside = st.number_input("Overnight Stays (Outside M25)", 0)
    miles = st.number_input("Mileage (miles)", 0)
    overnight_inside = st.number_input("Overnight Stays (Inside M25)", 0)
    flights_cost = st.number_input("Flights / Rail (£)", 0.0)
    other_cost = st.number_input("Other Costs (£)", 0.0)
    discount = st.number_input("Discount (£)", 0.0)

    # Ensure required data exists
    if "office_selling" in rates and "site_selling" in rates:

            # ✅ Use selling rates (NOT margin formula)
        office_rate = rates.get("office_selling", 0)
        site_rate = rates.get("site_selling", 0)
        
       
        labour_total = (
            (office_day * office_rate) +
            (site_day * site_rate) +
        
            (office_evening * office_rate) +
            (site_evening * site_rate) +
        
            (office_weekend * office_rate) +
            (site_weekend * site_rate)
)
    
        # ✅ Peer review (10% of office hours)
        peer_review = 0.1 * office_day * office_rate
        
        labour_total += peer_review

        # Expenses
        # ==========================
        # Expenses
        # ==========================
        
        expenses_total = (
            overnight_outside * rates.get("outside_m25", 0)
            + overnight_inside * rates.get("inside_m25", 0)
            + miles * rates.get("mileage", 0)
            + flights_cost * 1.15
        )
        
        # ==========================
        # Other Costs
        # ==========================
        
        other_cost_selling = other_cost * 1.2
        
        # ==========================
        # Final Total
        # ==========================
        
        subtotal = labour_total + expenses_total + other_cost_selling
        total_price = subtotal - discount

        # Display
        st.markdown("### Breakdown")

        
        st.write(f"Labour: £{labour_total:,.2f}")
        st.write(f"Expenses: £{expenses_total:,.2f}")
        st.write(f"Other Costs: £{other_cost_selling:,.2f}")
        st.write(f"Discount: -£{discount:,.2f}")
        
        st.markdown("---")
        st.metric("Total Price", f"£{total_price:,.2f}")

    else:
        st.error("⚠️ Could not extract rates from template")

# ==========================
# WORKS INPUT
# ==========================
st.subheader("🛠️ Works & Pricing")

col1, col2 = st.columns(2)

with col1:
    work_description = st.text_input("Work Description")

with col2:
    work_price = st.text_input("Price (£)")

if st.button("➕ Add Work"):
    if work_description and work_price:
        try:
            price = float(work_price.replace(",", ""))
            st.session_state.works_list.append({
                "description": work_description,
                "price": price
            })
            st.success("Work added")
        except:
            st.error("Enter a valid price")

if st.session_state.works_list:
    st.markdown("### Current Works")
    for i, work in enumerate(st.session_state.works_list):
        st.write(f"{i+1}. {work['description']} - £{work['price']:,.2f}")

st.markdown("---")

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
