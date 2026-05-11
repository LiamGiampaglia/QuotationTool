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

    # ✅ Site Name with guidance
    site_name = st.text_input(
        "Site Name",
        placeholder="e.g. Coventry, London and Warrington"
    )
    st.caption("If multiple sites, use format: Coventry, London and Warrington")

with col2:
    project_name = st.text_input("Project Name")

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
        st.caption("If multiple types, use format: HGV and Grey Fleet")

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

# Display current works
if st.session_state.works_list:
    st.markdown("### Current Works")
    for i, work in enumerate(st.session_state.works_list):
        st.write(f"{i+1}. {work['description']} - £{work['price']:,.2f}")

st.markdown("---")

# ==========================
# PLACEHOLDER REPLACEMENT
# ==========================
def replace_placeholders(doc, data):

    def replace_in_runs(paragraph):
        for run in paragraph.runs:
            for key, value in data.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, str(value))

    def replace_split_placeholder(paragraph):
        full_text = "".join(run.text for run in paragraph.runs)

        updated_text = full_text
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"
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

# ==========================
# TABLE POPULATION
# ==========================
def fill_works_table(doc, works_list):

    if not doc.tables:
        st.error("❌ No tables found in document")
        return doc

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
    last_row = table.rows[-1].cells
    last_row[2].text = f"£{total:,.2f}"

    return doc

# ==========================
# GENERATE DOCUMENT
# ==========================
if st.button("📄 Generate Word Document"):

    if not customer_name or not project_name:
        st.error("Customer Name and Project Name are required.")
    else:
        todays_date = datetime.now().strftime("%d %B %Y")

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
                "TodaysDate": todays_date
            }

            doc = replace_placeholders(doc, data)

            if st.session_state.works_list:
                doc = fill_works_table(doc, st.session_state.works_list)

            output = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            doc.save(output.name)

            st.success("✅ Document generated successfully!")

            with open(output.name, "rb") as f:
                st.download_button(
                    "⬇ Download Quote",
                    f,
                    file_name=f"{customer_name}_quote.docx"
                )

# ==========================
# INFO SECTION
# ==========================
with st.expander("📘 App Info"):
    st.write("Templates are automatically selected based on the chosen quote type.")
