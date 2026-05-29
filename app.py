from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
import streamlit as st
from docx import Document
import tempfile
from datetime import datetime

# ==========================
# SESSION STATE
# ==========================
if "works_list" not in st.session_state:
    st.session_state.works_list = []

# ==========================
# PAGE SETUP
# ==========================
st.set_page_config(page_title="Energy Quote Tool", layout="centered")
st.title("⚡ Energy Quote Generator")

# ==========================
# QUOTE TYPE
# ==========================
quote_type = st.selectbox(
    "Select Quote Type",
    [
        "Energy Efficiency Audit",
        "Metering Assessment",
        "EE and Metering",
        "ESOS P4",
        "ESOS P4 and Transport",
        "ESOS P4 Transport",
    ]
)

# ==========================
# TEMPLATE MAPPING
# ==========================
TEMPLATE_MAP = {
    "Energy Efficiency Audit": "templates/Energy Efficiency Audit Template.docx",
    "Metering Assessment": "templates/Metering Assessment Template.docx",
    "EE and Metering": "templates/EE Audit and Metering Template.docx",
    "ESOS P4": "templates/ESOS P4 Template.docx",
    "ESOS P4 and Transport": "templates/ESOS P4 and Transport Template.docx",
    "ESOS P4 Transport": "templates/ESOS P4 Transport Template.docx"
}

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

    def replace_in_runs(paragraph):
        for run in paragraph.runs:
            for key, value in data.items():
                placeholder = f"{{{{{key}}}}}"
                
                if placeholder == "{{PaymentTerms}}":
                    continue  # skip payment placeholder

                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, str(value))



    def replace_split_placeholder(paragraph):
        full_text = "".join(run.text for run in paragraph.runs)
    
        updated_text = full_text
    
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"
    
            # ✅ Skip PaymentTerms
            if placeholder == "{{PaymentTerms}}":
                continue
    
            updated_text = updated_text.replace(placeholder, str(value))
    
        if updated_text != full_text:
            if len(paragraph.text) < 300:
                paragraph.runs[0].text = updated_text
                for i in range(1, len(paragraph.runs)):
                    paragraph.runs[i].text = ""


    for paragraph in doc.paragraphs:
        replace_in_runs(paragraph)
        replace_split_placeholder(paragraph)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_runs(paragraph)
                    replace_split_placeholder(paragraph)

    return doc



def fill_works_table(doc, works_list):

    target_table = None

    # ✅ Find the table with "Work Description" header
    for table in doc.tables:
        first_row = table.rows[0].cells

        if "Work Description" in first_row[1].text:
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

    if not customer_name or not project_name:
        st.error("Customer Name and Project Name are required.")

    elif not st.session_state.works_list:
        st.error("Please add at least one work item before generating the document.")

    else:
        template_path = TEMPLATE_MAP.get(quote_type)
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
