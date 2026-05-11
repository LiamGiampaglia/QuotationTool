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
    st.caption("If multiple sites, use: Coventry, London and Warrington")

with col2:
    project_name = st.text_input("Project Name")

# ==========================
# 📄 FILE NAMING SECTION
# ==========================
st.markdown("---")
st.subheader("📄 File Naming")

# Document Type options
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

# Subject options
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

    document_type_label = st.selectbox(
        "Document Type",
        list(document_type_options.keys())
    )
    document_type = document_type_options[document_type_label]

    subject_label = st.selectbox(
        "Subject",
        list(subject_options.keys())
    )
    subject = subject_options[subject_label]

with col2:
    unique_id = st.selectbox(
        "Unique Identifier",
        [f"{i:02d}" for i in range(1, 21)]
    )

    revision_code_label = st.selectbox(
        "Revision Code",
        ["Contractual (External)", "Preliminary (Internal)"]
    )
    revision_code = "C" if "Contractual" in revision_code_label else "P"

    revision_number = st.selectbox(
        "Revision Number",
        [f"{i:02d}" for i in range(1, 21)]
    )

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
# FUNCTIONS
# ==========================
def replace_placeholders(doc, data):
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            paragraph.text = paragraph.text.replace(f"{{{{{key}}}}}", str(value))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in data.items():
                    cell.text = cell.text.replace(f"{{{{{key}}}}}", str(value))

    return doc


def fill_works_table(doc, works_list):
    table = doc.tables[0]
    start_row = 1

    for i, work in enumerate(works_list):
        if start_row + i >= len(table.rows) - 1:
            row_cells = table.add_row().cells
        else:
            row_cells = table.rows[start_row + i].cells

        row_cells[0].text = str(i + 1)
        row_cells[1].text = work["description"]
        row_cells[2].text = f"£{work['price']:,.2f}"

    total = sum(work["price"] for work in works_list)
    table.rows[-1].cells[2].text = f"£{total:,.2f}"

    return doc

# ==========================
# GENERATE DOCUMENT
# ==========================
if st.button("📄 Generate Word Document"):

    if not customer_name or not project_name:
        st.error("Customer Name and Project Name are required.")
    else:

        template_path = TEMPLATE_MAP.get(quote_type)

        if not template_path:
            st.error("Template not found.")
        else:

            doc = Document(template_path)

            data = {
                "CustomerName": customer_name,
                "NumberOfSites": number_of_sites,
                "SiteName": site_name,
                "NumberOfTransport": number_of_transport,
                "TransportType": transport_type,
                "ProjectName": project_name,
                "TodaysDate": datetime.now().strftime("%d %B %Y")
            }

            doc = replace_placeholders(doc, data)

            if st.session_state.works_list:
                doc = fill_works_table(doc, st.session_state.works_list)

            # ✅ Filename generation
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
