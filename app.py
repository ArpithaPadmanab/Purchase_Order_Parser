import streamlit as st
import imaplib
import email
from email.header import decode_header
import datetime
import pdfplumber
import pandas as pd
import io

def parse_pdf_fields(file_data):
    extracted_data = {
        "Vendor Number": None,
        "Purchase Order Number": None,
        "Vendor Address": None,
        "Item Description": [],
        "Quantity": [],
        "Unit Price": [],
        "Net Price": [],
        "GST": None,
        "Total Value": None,
    }

    with pdfplumber.open(io.BytesIO(file_data)) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"

        lines = full_text.splitlines()
        for line in lines:
            if "Vendor Number" in line:
                extracted_data["Vendor Number"] = line.split(":")[-1].strip()
            elif "Purchase Order" in line:
                extracted_data["Purchase Order Number"] = line.split(":")[-1].strip()
            elif "Vendor Address" in line:
                extracted_data["Vendor Address"] = line.split(":")[-1].strip()
            elif "GST" in line:
                extracted_data["GST"] = line.split(":")[-1].strip()
            elif "Total Value" in line or "Grand Total" in line:
                extracted_data["Total Value"] = line.split(":")[-1].strip()

        for page in pdf.pages:
            table = page.extract_table()
            if table:
                headers = [h.lower() if h else "" for h in table[0]]
                for row in table[1:]:
                    row_dict = dict(zip(headers, row))
                    extracted_data["Item Description"].append(row_dict.get("description", ""))
                    extracted_data["Quantity"].append(row_dict.get("quantity", ""))
                    extracted_data["Unit Price"].append(row_dict.get("unit price", ""))
                    extracted_data["Net Price"].append(row_dict.get("net price", ""))


    return extracted_data

def fetch_and_parse_po(email_user, email_pass, start_date, end_date):
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(email_user, email_pass)
    imap.select("INBOX")

    since = start_date.strftime("%d-%b-%Y")
    before = (end_date + datetime.timedelta(days=1)).strftime("%d-%b-%Y")

    status, msg_nums = imap.search(None, f'(SINCE {since} BEFORE {before})')
    results = []

    if status != "OK":
        return []

    for num in msg_nums[0].split():
        res, msg_data = imap.fetch(num, "(RFC822)")
        if res != "OK":
            continue
        msg = email.message_from_bytes(msg_data[0][1])

        for part in msg.walk():
            content_disposition = str(part.get("Content-Disposition", ""))
            filename = part.get_filename()
            if (
                filename
                and filename.lower().endswith(".pdf")
                and "purchase order" in filename.lower()
                and "attachment" in content_disposition
            ):
                file_data = part.get_payload(decode=True)
                parsed_data = parse_pdf_fields(file_data)
                parsed_data["Email Subject"] = msg["Subject"]
                parsed_data["Email From"] = msg["From"]
                parsed_data["Date"] = msg["Date"]
                results.append(parsed_data)

    imap.logout()
    return results

def flatten_results(results):
    flat_data = []
    for entry in results:
        max_len = max(len(entry["Item Description"]), 1)
        for i in range(max_len):
            row = {
                "Email From": entry["Email From"],
                "Email Subject": entry["Email Subject"],
                "Date": entry["Date"],
                "Vendor Number": entry["Vendor Number"],
                "Purchase Order Number": entry["Purchase Order Number"],
                "Vendor Address": entry["Vendor Address"],
                "Item Description": entry["Item Description"][i] if i < len(entry["Item Description"]) else "",
                "Quantity": entry["Quantity"][i] if i < len(entry["Quantity"]) else "",
                "Unit Price": entry["Unit Price"][i] if i < len(entry["Unit Price"]) else "",
                "Net Price": entry["Net Price"][i] if i < len(entry["Net Price"]) else "",
                "GST": entry["GST"],
                "Total Value": entry["Total Value"],
            }
            flat_data.append(row)
    return pd.DataFrame(flat_data)

# === Streamlit App ===
st.set_page_config(page_title="Purchase Order Extractor", layout="wide")
st.title("📧 Purchase Order Email Extractor (by PDF Filename)")

with st.sidebar:
    st.header("🔐 Email Credentials & Date Filter")
    email_id = st.text_input("📧 Gmail Address", value="")
    app_password = st.text_input("🔑 App Password", type="password")
    start_date = st.date_input("📆 Start Date", datetime.date.today() - datetime.timedelta(days=30))
    end_date = st.date_input("📆 End Date", datetime.date.today())

    run = st.button("🚀 Fetch Purchase Orders")

if run:
    if not email_id or not app_password:
        st.warning("Please enter both Gmail address and App Password.")
    else:
        with st.spinner("🔍 Connecting to Gmail and scanning emails..."):
            results = fetch_and_parse_po(email_id, app_password, start_date, end_date)
            if results:
                df = flatten_results(results)
                st.success(f"✅ Found {len(results)} emails with purchase order PDFs!")
                st.dataframe(df)

                output = io.BytesIO()
                df.to_excel(output, index=False, engine="openpyxl")

                st.download_button(
                    label="📥 Download Excel",
                    data=output.getvalue(),
                    file_name="purchase_order_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No emails with 'purchase order' in PDF filename found in the selected range.")