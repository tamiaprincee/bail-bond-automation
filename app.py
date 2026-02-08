import os
import streamlit as st
from docx import Document
from datetime import datetime
import smtplib
from email.message import EmailMessage

# === Constants ===
TEMPLATE_FILE = "Surety_Bond_Template.docx"
OUTPUT_DIR = "outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# === Fill Template ===
def fill_template(template_path, output_path, data):
    doc = Document(template_path)

    for para in doc.paragraphs:
        for run in para.runs:
            for key, value in data.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)
                    run.bold = True  # Bold the inserted value

    doc.save(output_path)

# === Email Attachment ===
def send_email_with_attachment(receiver_email, file_path):
    sender_email = "BigDawgBailBondz@gmail.com"
    app_password = "ebny jvku rexi zkiw"  # Replace with your real app password

    msg = EmailMessage()
    msg['Subject'] = "New Surety Bond"
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg.set_content("Please find the completed bond form attached.")

    with open(file_path, 'rb') as f:
        msg.add_attachment(
            f.read(),
            maintype='application',
            subtype='octet-stream',
            filename=os.path.basename(file_path)
        )

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(sender_email, app_password)
        smtp.send_message(msg)

# === Streamlit UI ===
st.title("📝 Bail Bond Form Entry")

with st.form("bond_form"):
    data = {
        "PRINCIPAL": st.text_input("Principal Name"),
        "SID": st.text_input("SID Number"),
        "DOB": st.text_input("Date of Birth (MM/DD/YYYY)"),
        "SEX": st.text_input("Sex"),
        "RACE": st.text_input("Race"),
        "CASE": st.text_input("Case Number"),
        "CHARGE": st.text_input("Charge"),
        "MISDEMEANOR_OR_FELONY": st.text_input("Misdemeanor or Felony"),
        "ADDRESS": st.text_input("Mailing Address"),
        "DL": st.text_input("Driver’s License Number"),
        "STATE": st.text_input("DL State"),
        "HT": st.text_input("Height"),
        "WT": st.text_input("Weight"),
        "HAIR": st.text_input("Hair Color"),
        "EYES": st.text_input("Eye Color"),
        "COUNTY": st.text_input("County"),
        "BOND": st.text_input("Bond"),
        "BACKING_AMOUNT": st.text_input("Amount to Back the Bond - Numeric"),
        "BACKING_AMOUNT_TEXT": st.text_input("Amount to Back the Bond - Text"),
        "SIGNED_AND_DATED": st.text_input("Signed and Dated (MM/DD/YYYY)")
    }
    submitted = st.form_submit_button("Generate Bond Form")

if submitted:
    # 1. Get the current year automatically
    current_year = str(datetime.now().year)
    
    # 2. Add it to the data dictionary so the 'fill_template' function finds it
    data["CURRENT_YEAR"] = current_year

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

