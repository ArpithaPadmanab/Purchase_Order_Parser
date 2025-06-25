import streamlit as st
import imaplib
import email
from email.header import decode_header
import datetime
import pdfplumber
import pandas as pd
import io

def parse_pdf_fields(file_data):
    import re
    extracted_data = {
        "Vendor Number": None,
        "PO Number": None,
        "PO Date": None,
        "Vendor Address": None,
        "Bill To": None,
        "Ship To": None,
        "Item Description": [],
        "Quantity": [],
        "Unit": [],
        "Unit Price": [],
        "Net Price": [],
        "State GST": None,
        "Central GST": None,
        "Total Value": None,
    }

    with pdfplumber.open(io.BytesIO(file_data)) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"

    # Safe regex extraction
    match = re.search(r"Vendor No\.\s*(\d+)", text)
    if match:
        extracted_data["Vendor Number"] = match.group(1)

    match = re.search(r"PO Number\s*(\d+)", text)
    if match:
        extracted_data["PO Number"] = match.group(1)

    match = re.search(r"PO Date\s*([\d.]+)", text)
    if match:
        extracted_data["PO Date"] = match.group(1)

    match = re.search(r"State GST\s*\n?\s*9%\s*\n?\s*([\d,.]+)", text)
    if match:
        extracted_data["State GST"] = match.group(1)

    match = re.search(r"Central GST\s*\n?\s*9%\s*\n?\s*([\d,.]+)", text)
    if match:
        extracted_data["Central GST"] = match.group(1)

    match = re.search(r"Total Order Value \( INR \)\s*([\d,\.]+)", text)
    if match:
        extracted_data["Total Value"] = match.group(1)

    # Extract Vendor Address
    vendor_addr_match = re.search(r"Vendor\s+(.*?)\s+Buyers Name:", text, re.DOTALL)
    if vendor_addr_match:
        extracted_data["Vendor Address"] = vendor_addr_match.group(1).strip()

    # Extract Bill To and Ship To
    bill_to_match = re.search(r"Bill To \(Invoice To\)\s+(.*?)\s+Ship To", text, re.DOTALL)
    ship_to_match = re.search(r"Ship To \(Deliver To\)\s+(.*?)\s+Vendor", text, re.DOTALL)
    if bill_to_match:
        extracted_data["Bill To"] = bill_to_match.group(1).strip()
    if ship_to_match:
        extracted_data["Ship To"] = ship_to_match.group(1).strip()

    # Extract item table block (known format from your sample)
    item_match = re.search(r"Item Material Code/.*?Qty\.\s+Unit\s+Deliv\. Date.*?\n(.*?)\n\s+State GST", text, re.DOTALL)
    if item_match:
        item_block = item_match.group(1).strip()
        item_lines = item_block.splitlines()
        for line in item_lines:
            parts = re.split(r'\s{2,}', line.strip())
            if len(parts) >= 4:
                extracted_data["Quantity"].append(parts[0])
                extracted_data["Unit"].append(parts[1])
                extracted_data["Unit Price"].append(parts[2].replace(",", ""))
                extracted_data["Net Price"].append(parts[3].replace(",", ""))
                extracted_data["Item Description"].append("Travel Charges (PO#...)")  # You can improve this

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
                "Email From": entry.get("Email From", ""),
                "Email Subject": entry.get("Email Subject", ""),
                "Date": entry.get("Date", ""),
                "Vendor Number": entry.get("Vendor No.", ""),
                "PO Number": entry.get("PO Number", ""),
                "PO Date": entry.get("PO Date", ""),
                "Vendor Address": entry.get("Vendor", ""),
                "Bill To": entry.get("Bill To", ""),
                "Ship To": entry.get("Ship To", ""),
                "Item Description": entry["Item Description"][i] if i < len(entry["Item Description"]) else "",
                "Quantity": entry["Quantity"][i] if i < len(entry["Quantity"]) else "",
                "Unit": entry["Unit"][i] if i < len(entry["Unit"]) else "",
                "Unit Price": entry["Unit Price"][i] if i < len(entry["Unit Price"]) else "",
                "Net Price": entry["Net Price"][i] if i < len(entry["Net Price"]) else "",
                "State GST": entry.get("State GST", ""),
                "Central GST": entry.get("Central GST", ""),
                "Total Order Value": entry.get("Total Value", "")
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