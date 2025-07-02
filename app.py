import os
import streamlit as st
from docx import Document
from datetime import datetime
import subprocess
import smtplib
from email.message import EmailMessage
import PyPDF2  # For PDF reading

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
TEMPLATE_FILE = "Surety_Bond_Template.docx"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Convert .doc to .docx
def convert_doc_to_docx(input_path):
    subprocess.run([
        "soffice", "--headless", "--convert-to", "docx", "--outdir", UPLOAD_DIR, input_path
    ], check=True)
    base = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(UPLOAD_DIR, base + ".docx")

# Extract text from PDF
def extract_text_from_pdf(pdf_path):
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        return text.splitlines()

# Parse lines of text and extract data
def extract_data_from_lines(lines):
    data = {}
    for line in lines:
        line = line.strip()
        if "SID" in line: data['SID'] = line.split()[-1]
        elif "DOB" in line: data['DOB'] = line.split()[-1]
        elif "SEX" in line: data['SEX'] = line.split()[-1]
        elif "RACE" in line: data['RACE'] = line.split()[-1]
        elif "CASE" in line.upper(): data['CASE'] = line.split()[-1]
        elif "Charge" in line: data['CHARGE'] = ' '.join(line.split()[1:])
        elif "Offense" in line: data['OFFENSE'] = line.split(":")[-1].strip()
        elif "Principal" in line: data['PRINCIPAL'] = line.split(":")[-1].strip()
        elif "Address" in line: data['ADDRESS'] = line.split(":")[-1].strip()
        elif "DL" in line: data['DL'] = line.split()[-1]
        elif "STATE" in line: data['STATE'] = line.split()[-1]
        elif "HT" in line: data['HT'] = line.split()[-1]
        elif "WT" in line: data['WT'] = line.split()[-1]
        elif "HAIR" in line: data['HAIR'] = line.split()[-1]
        elif "EYES" in line: data['EYES'] = line.split()[-1]
        elif "SUM" in line.upper(): data['SUM'] = line.split()[-1]
        elif "County" in line: data['COUNTY'] = line.split(":")[-1].strip()
        elif "Signed and Dated" in line or "Date Signed" in line: data['SIGNED_AND_DATED'] = line.split(":")[-1].strip()
    return data

# Wrapper to handle both docx and pdf
def extract_data(doc_path):
    ext = os.path.splitext(doc_path)[-1].lower()
    if ext == ".pdf":
        lines = extract_text_from_pdf(doc_path)
    else:
        doc = Document(doc_path)
        lines = [para.text for para in doc.paragraphs]
    return extract_data_from_lines(lines)

# Fill Word template
def fill_template(template_path, output_path, data):
    doc = Document(template_path)
    for para in doc.paragraphs:
        for key, value in data.items():
            if f"{{{{{key}}}}}" in para.text:
                para.text = para.text.replace(f"{{{{{key}}}}}", value)
    doc.save(output_path)

# Email bond form
def send_email_with_attachment(receiver_email, file_path, subject="New Surety Bond", body="Please find the completed bond form attached."):
    sender_email = "BigDawgBailBondz@gmail.com"
    app_password = "kuyb gdxu llhg nzou"

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg.set_content(body)

    with open(file_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=os.path.basename(file_path))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(sender_email, app_password)
        smtp.send_message(msg)

# Streamlit UI
st.title("📝 Bail Bond Form Automation")

uploaded_file = st.file_uploader("Upload Jail Form (.doc, .docx, or .pdf)", type=["doc", "docx", "pdf"])
if uploaded_file is not None:
    raw_path = os.path.join(UPLOAD_DIR, uploaded_file.name)
    with open(raw_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    ext = uploaded_file.name.lower().split(".")[-1]
    if ext == "doc":
        st.info("Converting .doc to .docx...")
        docx_path = convert_doc_to_docx(raw_path)
    else:
        docx_path = raw_path

    st.success("✅ File ready! Extracting information...")

    try:
        data = extract_data(docx_path)
        st.write("📄 Extracted Info:", data)

        if st.button("Generate Bond Form"):
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"filled_surety_bond_{timestamp}.docx"
            output_path = os.path.join(OUTPUT_DIR, output_file)
            fill_template(TEMPLATE_FILE, output_path, data)

            with open(output_path, "rb") as f:
                st.download_button("⬇️ Download Completed Bond Form", f, file_name=output_file)

            try:
                send_email_with_attachment("3gtexan@gmail.com", output_path)
                st.success("📧 A copy has been emailed to 3gtexan@gmail.com")
            except Exception as e:
                st.warning(f"⚠️ Document created, but email failed to send: {e}")
    except Exception as e:
        st.error(f"⚠️ Error during processing: {e}")
