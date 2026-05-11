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

st.markdown("---")

# ==========================
# TEMPLATE UPLOAD
# ==========================
uploaded_file = st.file_uploader(
    "Upload Template (.docx)", type=["docx"]
)

st.markdown("---")

# ==========================
# INPUT FIELDS
# ==========================
col1, col2 = st.columns(2)

with col1:
    customer_name = st.text_input("Customer Name")
    number_of_sites = st.text_input("Number of Sites")
    site_name = st.text_input("Site Name")

with col2:
    project_name = st.text_input("Project Name")

# ==========================
# ✅ TRANSPORT SECTION
# ==========================
st.markdown("---")

if "Transport" in quote_type:
    st.subheader("🚚 Transport Details")

    col1, col2 = st.columns(2)

    with col1:
        number_of_transport = st.text_input("Number of Transports")

    with col2:
        transport_type = st.text_input("Transport Type")

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

# Display works
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

    # Total row
    total = sum(work["price"] for work in works_list)
    last_row = table.rows[-1].cells
    last_row[2].text = f"£{total:,.2f}"

    return doc

# ==========================
# GENERATE DOCUMENT
# ==========================
if st.button("📄 Generate Word Document"):

    if uploaded_file is None:
        st.error("Please upload a template.")
    elif not customer_name or not project_name:
        st.error("Customer Name and Project Name are required.")
    else:
        todays_date = datetime.now().strftime("%d %B %Y")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uploaded_file.read())
            template_path = tmp.name

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
# HELP
# ==========================
with st.expander("📘 Template Help"):
    st.write("""
✅ PLACEHOLDERS:
{{CustomerName}}
{{NumberOfSites}}
{{SiteName}}
{{NumberOfTransport}}
{{TransportType}}
{{ProjectName}}
{{TodaysDate}}

✅ TABLE STRUCTURE:

| Item | Work Description | Price |
|------|-----------------|------|
| 1    | Placeholder     | Placeholder |
| ...  |                 |      |
|      | Total (ex. VAT) |      |

✅ NOTES:
- Table must be a normal Word table
- Not inside shapes/textboxes
- Blue shapes must be "Behind Text"
""")
