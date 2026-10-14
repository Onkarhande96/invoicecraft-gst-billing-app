"""
InvoiceCraft — Automated GST Invoice & Billing Generator SaaS
Engineered by Onkar Hande (BCA / MCA)
Streamlit Cloud & Enterprise Production Ready
"""

import io
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="InvoiceCraft — GST Billing & Invoice SaaS",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==========================================
# INDIAN STATES & GST CODES MAPPING
# ==========================================
INDIAN_STATES = [
    "27 - Maharashtra",
    "29 - Karnataka",
    "07 - Delhi",
    "24 - Gujarat",
    "33 - Tamil Nadu",
    "36 - Telangana",
    "09 - Uttar Pradesh",
    "19 - West Bengal",
    "06 - Haryana",
    "08 - Rajasthan",
    "32 - Kerala",
    "23 - Madhya Pradesh",
    "10 - Bihar",
    "21 - Odisha",
    "03 - Punjab",
    "30 - Goa",
    "99 - Other / Export",
]

# Number to Indian Words Converter (Pure Python)
def number_to_words_inr(n: float) -> str:
    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
            "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
            "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

    def convert_below_thousand(num):
        res = ""
        if num >= 100:
            res += ones[num // 100] + " Hundred "
            num %= 100
        if num >= 20:
            res += tens[num // 10] + " "
            num %= 10
        if num > 0:
            res += ones[num] + " "
        return res.strip()

    num_int = int(round(n))
    if num_int == 0:
        return "Rupees Zero Only"

    parts = []
    # Crores
    crores = num_int // 10000000
    if crores > 0:
        parts.append(convert_below_thousand(crores) + " Crore")
        num_int %= 10000000

    # Lakhs
    lakhs = num_int // 100000
    if lakhs > 0:
        parts.append(convert_below_thousand(lakhs) + " Lakh")
        num_int %= 100000

    # Thousands
    thousands = num_int // 1000
    if thousands > 0:
        parts.append(convert_below_thousand(thousands) + " Thousand")
        num_int %= 1000

    # Hundreds & below
    if num_int > 0:
        parts.append(convert_below_thousand(num_int))

    return "Rupees " + " ".join(parts).strip() + " Only"

# ==========================================
# PRESET TEMPLATES
# ==========================================
PRESETS = {
    "IT Consulting & Data Automation (18% GST)": {
        "seller_name": "Apex Data & Tech Solutions",
        "seller_gstin": "27AAACA1234A1Z5",
        "seller_pan": "AAACA1234A",
        "seller_address": "402 Tech Park, Baner, Pune, Maharashtra - 411045",
        "seller_phone": "+91 98765 43210",
        "seller_email": "billing@apextechsolutions.in",
        "seller_state": "27 - Maharashtra",
        "client_name": "Zenith Retail Global Pvt Ltd",
        "client_gstin": "27BBBCB5678B2Z9",
        "client_address": "12, Business Bay, Andheri East, Mumbai - 400069",
        "client_state": "27 - Maharashtra",
        "items": [
            {"description": "Automated n8n Inbound Lead Routing Workflow", "hsn": "998314", "qty": 1.0, "unit": "Nos", "rate": 35000.0, "discount": 0.0, "gst_rate": 18.0},
            {"description": "Multi-Depot Inventory Excel Cleanup & Reconciliation", "hsn": "998313", "qty": 2.0, "unit": "Days", "rate": 12000.0, "discount": 10.0, "gst_rate": 18.0},
            {"description": "Executive KPI Sales Analytics Dashboard Setup", "hsn": "998314", "qty": 1.0, "unit": "Nos", "rate": 25000.0, "discount": 0.0, "gst_rate": 18.0},
        ],
        "bank_name": "HDFC Bank Ltd",
        "bank_account": "50200012345678",
        "bank_ifsc": "HDFC0001234",
        "bank_branch": "Baner Pune Branch",
        "upi_id": "apextech@hdfcbank",
        "terms": "1. 50% advance, balance upon project delivery.\n2. Payment due within 15 days of invoice date.\n3. Late fee of 1.5% per month applicable after due date.",
    },
    "E-Commerce & Retail Supply (Intra-State)": {
        "seller_name": "Kulkarni Enterprise Mart",
        "seller_gstin": "27CCDCE9988C1Z2",
        "seller_pan": "CCDCE9988C",
        "seller_address": "Shop 14, Market Yard, Swargate, Pune - 411037",
        "seller_phone": "+91 98220 12345",
        "seller_email": "orders@kulkarnienterprise.com",
        "seller_state": "27 - Maharashtra",
        "client_name": "Deshmukh Agro Foods",
        "client_gstin": "27DDDDE1122D1Z8",
        "client_address": "Plot 88, MIDC Shiroli, Kolhapur - 416122",
        "client_state": "27 - Maharashtra",
        "items": [
            {"description": "Thermal Billing Rolls (80mm x 50m Box of 50)", "hsn": "481190", "qty": 10.0, "unit": "Box", "rate": 1250.0, "discount": 5.0, "gst_rate": 12.0},
            {"description": "Wireless Barcode Scanner Handheld USB", "hsn": "847190", "qty": 4.0, "unit": "Pcs", "rate": 2800.0, "discount": 0.0, "gst_rate": 18.0},
            {"description": "Electronic Weighing Scale (30Kg Certified)", "hsn": "842381", "qty": 2.0, "unit": "Nos", "rate": 4500.0, "discount": 8.0, "gst_rate": 18.0},
        ],
        "bank_name": "State Bank of India",
        "bank_account": "38920192837",
        "bank_ifsc": "SBIN0004567",
        "bank_branch": "Swargate Branch",
        "upi_id": "kulkarnimart@sbi",
        "terms": "1. Goods once sold will not be taken back without original packaging.\n2. Warranty handled by respective manufacturers.\n3. Subject to Pune jurisdiction.",
    },
    "Inter-State Digital Services (Maharashtra to Karnataka - IGST)": {
        "seller_name": "Hande Automation Studio",
        "seller_gstin": "27EEEFE3344E1Z4",
        "seller_pan": "EEEFE3344E",
        "seller_address": "Tech Enclave, Shivajinagar, Pune - 411005",
        "seller_phone": "+91 91987 65432",
        "seller_email": "hello@handeautomation.com",
        "seller_state": "27 - Maharashtra",
        "client_name": "Bangalore Cloud Infotech Pvt Ltd",
        "client_gstin": "29FFFGF5566F1Z1",
        "client_address": "Level 5, Embassy Tech Village, Outer Ring Rd, Bangalore - 560103",
        "client_state": "29 - Karnataka",
        "items": [
            {"description": "Custom API Webhook Integration for HubSpot CRM", "hsn": "998314", "qty": 1.0, "unit": "Nos", "rate": 48000.0, "discount": 0.0, "gst_rate": 18.0},
            {"description": "Automated Telegram Notification Bot Development", "hsn": "998314", "qty": 1.0, "unit": "Nos", "rate": 22000.0, "discount": 5.0, "gst_rate": 18.0},
        ],
        "bank_name": "ICICI Bank Ltd",
        "bank_account": "004505012345",
        "bank_ifsc": "ICIC0000045",
        "bank_branch": "Shivajinagar Pune",
        "upi_id": "handeauto@icici",
        "terms": "1. 100% Tax Invoice complying with IGST Interstate provisions.\n2. Payment terms: Net 10 days via NEFT/RTGS/UPI.\n3. Includes 30 days post-deployment support.",
    },
}

# ==========================================
# SESSION STATE INITIALIZATION
# ==========================================
if "current_preset" not in st.session_state:
    st.session_state.current_preset = "IT Consulting & Data Automation (18% GST)"

default_data = PRESETS[st.session_state.current_preset]

if "invoice_meta" not in st.session_state:
    st.session_state.invoice_meta = {
        "invoice_number": f"INV-{datetime.now().strftime('%Y')}-0089",
        "invoice_date": date.today(),
        "due_date": date.today(),
        "seller_name": default_data["seller_name"],
        "seller_gstin": default_data["seller_gstin"],
        "seller_pan": default_data["seller_pan"],
        "seller_address": default_data["seller_address"],
        "seller_phone": default_data["seller_phone"],
        "seller_email": default_data["seller_email"],
        "seller_state": default_data["seller_state"],
        "client_name": default_data["client_name"],
        "client_gstin": default_data["client_gstin"],
        "client_address": default_data["client_address"],
        "client_state": default_data["client_state"],
        "bank_name": default_data["bank_name"],
        "bank_account": default_data["bank_account"],
        "bank_ifsc": default_data["bank_ifsc"],
        "bank_branch": default_data["bank_branch"],
        "upi_id": default_data["upi_id"],
        "terms": default_data["terms"],
        "signatory": "Onkar Hande (Authorized Signatory)",
    }

if "invoice_items" not in st.session_state:
    st.session_state.invoice_items = list(default_data["items"])

# Function to load a preset
def load_preset(preset_key):
    p = PRESETS[preset_key]
    st.session_state.invoice_meta.update({
        "seller_name": p["seller_name"],
        "seller_gstin": p["seller_gstin"],
        "seller_pan": p["seller_pan"],
        "seller_address": p["seller_address"],
        "seller_phone": p["seller_phone"],
        "seller_email": p["seller_email"],
        "seller_state": p["seller_state"],
        "client_name": p["client_name"],
        "client_gstin": p["client_gstin"],
        "client_address": p["client_address"],
        "client_state": p["client_state"],
        "bank_name": p["bank_name"],
        "bank_account": p["bank_account"],
        "bank_ifsc": p["bank_ifsc"],
        "bank_branch": p["bank_branch"],
        "upi_id": p["upi_id"],
        "terms": p["terms"],
    })
    st.session_state.invoice_items = list(p["items"])

# ==========================================
# CUSTOM CSS FOR INVOICE STYLING
# ==========================================
st.markdown(
    """
    <style>
    /* Metric Cards */
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 14px 18px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.6rem !important;
        font-weight: 800 !important;
        color: #0f172a !important;
    }

    /* Print Preview Container */
    .invoice-box {
        max-width: 900px;
        margin: auto;
        padding: 30px 40px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.06);
        background: #ffffff;
        color: #1e293b;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        border-radius: 8px;
    }
    .inv-header {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        border-bottom: 2px solid #2563eb;
        padding-bottom: 16px;
        margin-bottom: 24px;
    }
    .inv-title {
        font-size: 26px;
        font-weight: 800;
        color: #2563eb;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .inv-badge {
        display: inline-block;
        background: #eff6ff;
        color: #1d4ed8;
        font-size: 12px;
        font-weight: 700;
        padding: 3px 10px;
        border-radius: 4px;
        margin-top: 4px;
    }
    .inv-grid-2 {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 30px;
        margin-bottom: 24px;
        font-size: 13.5px;
        line-height: 1.5;
    }
    .inv-party-title {
        font-weight: 700;
        font-size: 13px;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 6px;
        border-bottom: 1px solid #f1f5f9;
        padding-bottom: 4px;
    }
    .inv-table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 24px;
        font-size: 13px;
    }
    .inv-table th {
        background-color: #f8fafc;
        color: #334155;
        font-weight: 700;
        padding: 10px 12px;
        text-align: left;
        border-bottom: 2px solid #e2e8f0;
    }
    .inv-table td {
        padding: 10px 12px;
        border-bottom: 1px solid #f1f5f9;
        color: #1e293b;
    }
    .inv-table tr:hover td {
        background-color: #fafafa;
    }
    .inv-summary-grid {
        display: grid;
        grid-template-columns: 1.2fr 1fr;
        gap: 30px;
        margin-bottom: 24px;
        font-size: 13.5px;
    }
    .inv-total-row {
        display: flex;
        justify-content: space-between;
        padding: 6px 0;
        border-bottom: 1px solid #f1f5f9;
    }
    .inv-grand-total {
        display: flex;
        justify-content: space-between;
        padding: 12px 0;
        font-size: 18px;
        font-weight: 800;
        color: #0f172a;
        border-top: 2px solid #0f172a;
        border-bottom: 2px solid #0f172a;
        margin-top: 8px;
    }
    .words-box {
        background: #f8fafc;
        border-radius: 6px;
        padding: 10px 14px;
        font-size: 12.5px;
        color: #475569;
        font-style: italic;
        margin-bottom: 20px;
        border-left: 3px solid #2563eb;
    }
    .inv-footer {
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
        padding-top: 20px;
        border-top: 1px solid #e2e8f0;
        font-size: 12px;
        color: #64748b;
    }
    .sign-box {
        text-align: center;
        border-top: 1px solid #cbd5e1;
        width: 220px;
        padding-top: 6px;
        font-weight: 600;
        color: #1e293b;
    }

    /* Print media style */
    @media print {
        body * {
            visibility: hidden;
        }
        .invoice-box, .invoice-box * {
            visibility: visible;
        }
        .invoice-box {
            position: absolute;
            left: 0;
            top: 0;
            width: 100% !important;
            border: none !important;
            box-shadow: none !important;
            padding: 0 !important;
        }
        .no-print {
            display: none !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# SIDEBAR CONTROLS
# ==========================================
with st.sidebar:
    st.markdown("### 🧾 InvoiceCraft GST")
    st.caption("Automated GST Invoice & Billing SaaS")
    st.markdown("---")

    st.markdown("#### ⚡ Quick Load Business Presets")
    selected_preset = st.selectbox(
        "Choose Template Preset:",
        list(PRESETS.keys()),
        index=list(PRESETS.keys()).index(st.session_state.current_preset),
    )
    if st.button("Apply Template Preset ➔", use_container_width=True):
        st.session_state.current_preset = selected_preset
        load_preset(selected_preset)
        st.success(f"Loaded: {selected_preset}")
        st.rerun()

    st.markdown("---")
    st.markdown("#### ⚙️ Quick Invoice Settings")
    st.session_state.invoice_meta["invoice_number"] = st.text_input(
        "Invoice Number", st.session_state.invoice_meta["invoice_number"]
    )
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.session_state.invoice_meta["invoice_date"] = st.date_input(
            "Invoice Date", st.session_state.invoice_meta["invoice_date"]
        )
    with col_d2:
        st.session_state.invoice_meta["due_date"] = st.date_input(
            "Due Date", st.session_state.invoice_meta["due_date"]
        )

    st.markdown("---")
    st.markdown(
        """
        <div style="font-size: 0.8rem; color: #64748b;">
        <strong>InvoiceCraft Cloud SaaS</strong><br>
        Engineered by Onkar Hande (BCA / MCA)<br>
        <a href="https://onkarhande96.github.io" target="_blank" style="color: #2563eb; text-decoration: none;">Live Portfolio ↗</a>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ==========================================
# CALCULATIONS ENGINE
# ==========================================
seller_state_code = st.session_state.invoice_meta["seller_state"][:2]
client_state_code = st.session_state.invoice_meta["client_state"][:2]
is_interstate = seller_state_code != client_state_code

# Calculate line items
processed_items = []
total_taxable = 0.0
total_cgst = 0.0
total_sgst = 0.0
total_igst = 0.0

for idx, itm in enumerate(st.session_state.invoice_items):
    qty = float(itm.get("qty", 1.0))
    rate = float(itm.get("rate", 0.0))
    disc_pct = float(itm.get("discount", 0.0))
    gst_pct = float(itm.get("gst_rate", 18.0))

    base_amt = qty * rate
    disc_amt = base_amt * (disc_pct / 100.0)
    taxable_val = base_amt - disc_amt

    if is_interstate:
        cgst_amt = 0.0
        sgst_amt = 0.0
        igst_amt = taxable_val * (gst_pct / 100.0)
    else:
        cgst_amt = taxable_val * ((gst_pct / 2.0) / 100.0)
        sgst_amt = taxable_val * ((gst_pct / 2.0) / 100.0)
        igst_amt = 0.0

    line_total = taxable_val + cgst_amt + sgst_amt + igst_amt

    total_taxable += taxable_val
    total_cgst += cgst_amt
    total_sgst += sgst_amt
    total_igst += igst_amt

    processed_items.append({
        "index": idx + 1,
        "description": itm["description"],
        "hsn": itm["hsn"],
        "qty": qty,
        "unit": itm["unit"],
        "rate": rate,
        "discount": disc_pct,
        "taxable": taxable_val,
        "gst_rate": gst_pct,
        "cgst": cgst_amt,
        "sgst": sgst_amt,
        "igst": igst_amt,
        "total": line_total,
    })

total_tax = total_cgst + total_sgst + total_igst
gross_total = total_taxable + total_tax
round_off = round(gross_total) - gross_total
net_payable = round(gross_total)
amount_words = number_to_words_inr(net_payable)

# ==========================================
# MAIN INTERFACE TABS
# ==========================================
st.title("🧾 InvoiceCraft — Automated GST Billing Generator")
st.markdown("Create GST-compliant Tax Invoices with automated intra/inter-state tax splitting, dynamic line items, and audit-ready Excel & PDF downloads.")

# Quick KPI Metrics
k1, k2, k3, k4 = st.columns(4)
k1.metric("Taxable Subtotal", f"₹{total_taxable:,.2f}")
if is_interstate:
    k2.metric("IGST (Inter-State)", f"₹{total_igst:,.2f}")
else:
    k2.metric("CGST + SGST (Intra-State)", f"₹{(total_cgst + total_sgst):,.2f}")
k3.metric("Net Payable Total", f"₹{net_payable:,.0f}")
k4.metric("GST Type", "Inter-State (IGST)" if is_interstate else "Intra-State (CGST+SGST)")

st.markdown("<br>", unsafe_allow_html=True)

tab_preview, tab_edit_meta, tab_edit_items, tab_export = st.tabs([
    "👁️ Live Invoice Document Preview",
    "🏢 Business & Client Details",
    "📦 Line Items & Pricing (Add / Edit)",
    "📥 Downloads (Excel & PDF)",
])

# ------------------------------------------
# TAB 1: LIVE INVOICE DOCUMENT PREVIEW
# ------------------------------------------
with tab_preview:
    st.subheader("Tax Invoice Preview (Print Ready)")

    # Print / Browser PDF button using HTML JS
    col_p1, col_p2 = st.columns([3, 1])
    with col_p1:
        st.caption("Click 'Print / Save as PDF' below to save a vector PDF via your browser:")
    with col_p2:
        st.components.v1.html(
            """
            <button onclick="window.print()" style="
                background-color: #2563eb;
                color: white;
                border: none;
                padding: 9px 18px;
                border-radius: 6px;
                font-weight: 700;
                font-size: 14px;
                cursor: pointer;
                width: 100%;
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 8px;
            ">
                🖨️ Print / Save as PDF
            </button>
            """,
            height=45,
        )

    # Render HTML Invoice
    meta = st.session_state.invoice_meta
    rows_html = ""
    for item in processed_items:
        tax_col_content = f"IGST: {item['gst_rate']}% (₹{item['igst']:,.2f})" if is_interstate else f"CGST: {item['gst_rate']/2:.1f}%<br>SGST: {item['gst_rate']/2:.1f}%"
        rows_html += f"""
        <tr>
            <td style="text-align: center; color: #64748b;">{item['index']}</td>
            <td><strong>{item['description']}</strong></td>
            <td style="text-align: center; font-family: monospace;">{item['hsn']}</td>
            <td style="text-align: center;">{item['qty']} {item['unit']}</td>
            <td style="text-align: right;">₹{item['rate']:,.2f}</td>
            <td style="text-align: right;">{item['discount']}%</td>
            <td style="text-align: right; font-weight: 600;">₹{item['taxable']:,.2f}</td>
            <td style="text-align: right; font-size: 11px;">{tax_col_content}</td>
            <td style="text-align: right; font-weight: 700;">₹{item['total']:,.2f}</td>
        </tr>
        """

    tax_summary_html = ""
    if is_interstate:
        tax_summary_html = f"""
        <div class="inv-total-row">
            <span>Integrated GST (IGST):</span>
            <strong>₹{total_igst:,.2f}</strong>
        </div>
        """
    else:
        tax_summary_html = f"""
        <div class="inv-total-row">
            <span>Central GST (CGST):</span>
            <strong>₹{total_cgst:,.2f}</strong>
        </div>
        <div class="inv-total-row">
            <span>State GST (SGST):</span>
            <strong>₹{total_sgst:,.2f}</strong>
        </div>
        """

    invoice_html = f"""
    <div class="invoice-box">
        <div class="inv-header">
            <div>
                <h1 class="inv-title">{meta['seller_name']}</h1>
                <div style="font-size: 13px; color: #475569; margin-top: 4px;">{meta['seller_address']}</div>
                <div style="font-size: 12.5px; color: #475569; margin-top: 2px;">
                    <strong>GSTIN:</strong> {meta['seller_gstin']} | <strong>PAN:</strong> {meta['seller_pan']} | <strong>Phone:</strong> {meta['seller_phone']}
                </div>
            </div>
            <div style="text-align: right;">
                <div style="font-size: 22px; font-weight: 800; color: #0f172a;">TAX INVOICE</div>
                <span class="inv-badge">Original for Recipient</span>
                <div style="font-size: 13px; color: #334155; margin-top: 8px;">
                    <strong>Invoice #:</strong> {meta['invoice_number']}<br>
                    <strong>Date:</strong> {meta['invoice_date']}<br>
                    <strong>Due Date:</strong> {meta['due_date']}
                </div>
            </div>
        </div>

        <div class="inv-grid-2">
            <div>
                <div class="inv-party-title">Billed To (Client Details):</div>
                <div style="font-weight: 700; font-size: 15px; color: #0f172a;">{meta['client_name']}</div>
                <div style="color: #475569; margin: 3px 0;">{meta['client_address']}</div>
                <div><strong>GSTIN:</strong> {meta['client_gstin']}</div>
                <div><strong>State of Supply:</strong> {meta['client_state']}</div>
            </div>
            <div>
                <div class="inv-party-title">Dispatch / Supplier Location:</div>
                <div style="font-weight: 700; font-size: 15px; color: #0f172a;">{meta['seller_name']}</div>
                <div style="color: #475569; margin: 3px 0;">State of Origin: <strong>{meta['seller_state']}</strong></div>
                <div><strong>Place of Supply:</strong> {meta['client_state']}</div>
                <div style="color: #2563eb; font-weight: 600; margin-top: 4px;">
                    Tax Mechanism: {'Inter-State (IGST)' if is_interstate else 'Intra-State (CGST + SGST)'}
                </div>
            </div>
        </div>

        <table class="inv-table">
            <thead>
                <tr>
                    <th style="width: 5%; text-align: center;">#</th>
                    <th style="width: 32%;">Item Description</th>
                    <th style="width: 10%; text-align: center;">HSN/SAC</th>
                    <th style="width: 10%; text-align: center;">Qty</th>
                    <th style="width: 11%; text-align: right;">Rate (₹)</th>
                    <th style="width: 6%; text-align: right;">Disc %</th>
                    <th style="width: 11%; text-align: right;">Taxable (₹)</th>
                    <th style="width: 15%; text-align: right;">Tax Split</th>
                    <th style="width: 13%; text-align: right;">Total (₹)</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>

        <div class="inv-summary-grid">
            <div>
                <div class="inv-party-title">Bank & Payment Instructions:</div>
                <div style="font-size: 13px; line-height: 1.6; color: #334155;">
                    <strong>Bank Name:</strong> {meta['bank_name']}<br>
                    <strong>A/C No:</strong> {meta['bank_account']}<br>
                    <strong>IFSC Code:</strong> {meta['bank_ifsc']}<br>
                    <strong>Branch:</strong> {meta['bank_branch']}<br>
                    <strong>UPI ID:</strong> <span style="background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-family: monospace;">{meta['upi_id']}</span>
                </div>
            </div>
            <div>
                <div class="inv-total-row">
                    <span>Taxable Subtotal:</span>
                    <strong>₹{total_taxable:,.2f}</strong>
                </div>
                {tax_summary_html}
                <div class="inv-total-row">
                    <span>Round Off:</span>
                    <span>₹{round_off:+,.2f}</span>
                </div>
                <div class="inv-grand-total">
                    <span>Total Net Payable:</span>
                    <span>₹{net_payable:,.0f}</span>
                </div>
            </div>
        </div>

        <div class="words-box">
            <strong>Amount in Words:</strong> {amount_words}
        </div>

        <div class="inv-footer">
            <div style="max-width: 55%;">
                <strong>Terms & Conditions:</strong><br>
                <span style="white-space: pre-line; line-height: 1.4;">{meta['terms']}</span>
            </div>
            <div style="text-align: right;">
                <div style="font-size: 12px; margin-bottom: 45px;">For <strong>{meta['seller_name']}</strong></div>
                <div class="sign-box">{meta['signatory']}</div>
            </div>
        </div>
    </div>
    """
    st.markdown(invoice_html, unsafe_allow_html=True)

# ------------------------------------------
# TAB 2: BUSINESS & CLIENT DETAILS
# ------------------------------------------
with tab_edit_meta:
    st.subheader("🏢 Business (Seller) & Client (Buyer) Information")

    col_s1, col_s2 = st.columns(2)

    with col_s1:
        st.markdown("#### 🏪 Your Business Information")
        st.session_state.invoice_meta["seller_name"] = st.text_input("Business / Trade Name *", st.session_state.invoice_meta["seller_name"])
        st.session_state.invoice_meta["seller_gstin"] = st.text_input("GSTIN *", st.session_state.invoice_meta["seller_gstin"])
        st.session_state.invoice_meta["seller_pan"] = st.text_input("PAN Number", st.session_state.invoice_meta["seller_pan"])
        st.session_state.invoice_meta["seller_state"] = st.selectbox(
            "State of Registration *",
            INDIAN_STATES,
            index=INDIAN_STATES.index(st.session_state.invoice_meta["seller_state"]) if st.session_state.invoice_meta["seller_state"] in INDIAN_STATES else 0
        )
        st.session_state.invoice_meta["seller_address"] = st.text_area("Registered Address", st.session_state.invoice_meta["seller_address"], height=80)
        st.session_state.invoice_meta["seller_phone"] = st.text_input("Phone Number", st.session_state.invoice_meta["seller_phone"])
        st.session_state.invoice_meta["seller_email"] = st.text_input("Email Address", st.session_state.invoice_meta["seller_email"])

    with col_s2:
        st.markdown("#### 👤 Client / Customer Information")
        st.session_state.invoice_meta["client_name"] = st.text_input("Client / Company Name *", st.session_state.invoice_meta["client_name"])
        st.session_state.invoice_meta["client_gstin"] = st.text_input("Client GSTIN (Leave blank for unregistered)", st.session_state.invoice_meta["client_gstin"])
        st.session_state.invoice_meta["client_state"] = st.selectbox(
            "Place of Supply (State) *",
            INDIAN_STATES,
            index=INDIAN_STATES.index(st.session_state.invoice_meta["client_state"]) if st.session_state.invoice_meta["client_state"] in INDIAN_STATES else 0
        )
        st.session_state.invoice_meta["client_address"] = st.text_area("Client Billing Address", st.session_state.invoice_meta["client_address"], height=80)

        st.markdown("#### 🏦 Bank & Payment Details")
        b1, b2 = st.columns(2)
        with b1:
            st.session_state.invoice_meta["bank_name"] = st.text_input("Bank Name", st.session_state.invoice_meta["bank_name"])
            st.session_state.invoice_meta["bank_account"] = st.text_input("Account Number", st.session_state.invoice_meta["bank_account"])
        with b2:
            st.session_state.invoice_meta["bank_ifsc"] = st.text_input("IFSC Code", st.session_state.invoice_meta["bank_ifsc"])
            st.session_state.invoice_meta["upi_id"] = st.text_input("UPI ID (VPA)", st.session_state.invoice_meta["upi_id"])
        
        st.session_state.invoice_meta["bank_branch"] = st.text_input("Branch Name", st.session_state.invoice_meta["bank_branch"])

    st.markdown("---")
    st.markdown("#### 📜 Terms & Conditions & Signatory")
    t1, t2 = st.columns(2)
    with t1:
        st.session_state.invoice_meta["terms"] = st.text_area("Invoice Terms & Notes", st.session_state.invoice_meta["terms"], height=100)
    with t2:
        st.session_state.invoice_meta["signatory"] = st.text_input("Authorized Signatory Text", st.session_state.invoice_meta["signatory"])

    st.success("Business details updated! Check the '👁️ Live Invoice Document Preview' tab to see changes.")

# ------------------------------------------
# TAB 3: LINE ITEMS & PRICING
# ------------------------------------------
with tab_edit_items:
    st.subheader("📦 Invoice Line Items Management")
    st.markdown("Add, modify, or remove product items and services. Taxes and totals update in real-time.")

    # Show existing items
    for i, itm in enumerate(st.session_state.invoice_items):
        with st.expander(f"Item #{i+1}: {itm['description']} (₹{itm['rate']:,.2f} x {itm['qty']})", expanded=(i == 0)):
            c_desc, c_hsn, c_qty, c_unit = st.columns([3, 1.2, 1, 1])
            with c_desc:
                itm["description"] = st.text_input(f"Description #{i+1}", itm["description"], key=f"desc_{i}")
            with c_hsn:
                itm["hsn"] = st.text_input(f"HSN/SAC #{i+1}", itm["hsn"], key=f"hsn_{i}")
            with c_qty:
                itm["qty"] = st.number_input(f"Quantity #{i+1}", min_value=0.1, value=float(itm["qty"]), step=1.0, key=f"qty_{i}")
            with c_unit:
                itm["unit"] = st.selectbox(f"Unit #{i+1}", ["Nos", "Hrs", "Days", "Pcs", "Box", "Kgs", "Month"], index=["Nos", "Hrs", "Days", "Pcs", "Box", "Kgs", "Month"].index(itm.get("unit", "Nos")), key=f"unit_{i}")

            c_rate, c_disc, c_gst, c_del = st.columns([1.5, 1, 1.2, 1])
            with c_rate:
                itm["rate"] = st.number_input(f"Unit Rate (₹) #{i+1}", min_value=0.0, value=float(itm["rate"]), step=100.0, key=f"rate_{i}")
            with c_disc:
                itm["discount"] = st.number_input(f"Discount % #{i+1}", min_value=0.0, max_value=100.0, value=float(itm.get("discount", 0.0)), step=1.0, key=f"disc_{i}")
            with c_gst:
                gst_opts = [0.0, 5.0, 12.0, 18.0, 28.0]
                cur_gst = float(itm.get("gst_rate", 18.0))
                itm["gst_rate"] = st.selectbox(f"GST Rate % #{i+1}", gst_opts, index=gst_opts.index(cur_gst) if cur_gst in gst_opts else 3, key=f"gst_{i}")
            with c_del:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button(f"🗑️ Delete Item", key=f"del_{i}"):
                    if len(st.session_state.invoice_items) > 1:
                        st.session_state.invoice_items.pop(i)
                        st.rerun()
                    else:
                        st.warning("Invoice must contain at least 1 item.")

    st.markdown("---")
    st.markdown("#### ➕ Add New Item / Service")
    with st.form("add_item_form", clear_on_submit=True):
        a1, a2, a3, a4 = st.columns([3, 1.2, 1, 1])
        with a1:
            new_desc = st.text_input("Item Description *", placeholder="e.g. Python Script Automation")
        with a2:
            new_hsn = st.text_input("HSN/SAC Code", value="998314")
        with a3:
            new_qty = st.number_input("Quantity", min_value=0.1, value=1.0, step=1.0)
        with a4:
            new_unit = st.selectbox("Unit", ["Nos", "Hrs", "Days", "Pcs", "Box", "Kgs", "Month"])

        b1, b2, b3 = st.columns([1.5, 1, 1.2])
        with b1:
            new_rate = st.number_input("Unit Rate (₹) *", min_value=0.0, value=5000.0, step=500.0)
        with b2:
            new_disc = st.number_input("Discount %", min_value=0.0, max_value=100.0, value=0.0, step=1.0)
        with b3:
            new_gst = st.selectbox("GST Rate %", [0.0, 5.0, 12.0, 18.0, 28.0], index=3)

        add_btn = st.form_submit_button("➕ Add Item to Invoice", use_container_width=True)
        if add_btn:
            if not new_desc:
                st.error("Please enter item description.")
            else:
                st.session_state.invoice_items.append({
                    "description": new_desc.strip(),
                    "hsn": new_hsn.strip(),
                    "qty": new_qty,
                    "unit": new_unit,
                    "rate": new_rate,
                    "discount": new_disc,
                    "gst_rate": new_gst,
                })
                st.success(f"Added item: {new_desc}")
                st.rerun()

# ------------------------------------------
# TAB 4: DOWNLOADS & EXPORTS
# ------------------------------------------
with tab_export:
    st.subheader("📥 Export Audit-Ready Invoice Documents")
    st.markdown("Download this invoice in universal formats for accounting, audit compliance, or client dispatch.")

    d_col1, d_col2 = st.columns(2)

    with d_col1:
        st.markdown("#### 📊 Formatted Excel Invoice (.xlsx)")
        st.caption("Includes full company header, GSTIN details, dynamic line items, and audit-ready tax calculations.")

        # Build Excel Workbook using openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Tax_Invoice"

        # Styles
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        sub_font = Font(name="Calibri", size=10, italic=True)
        bold_font = Font(name="Calibri", size=11, bold=True)
        thin_border = Border(
            left=Side(style="thin", color="CBD5E1"),
            right=Side(style="thin", color="CBD5E1"),
            top=Side(style="thin", color="CBD5E1"),
            bottom=Side(style="thin", color="CBD5E1"),
        )

        # Company Header
        ws.merge_cells("A1:H1")
        ws["A1"] = meta["seller_name"]
        ws["A1"].font = header_font
        ws["A1"].fill = header_fill
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("A2:H2")
        ws["A2"] = f"{meta['seller_address']} | GSTIN: {meta['seller_gstin']} | State: {meta['seller_state']}"
        ws["A2"].font = sub_font
        ws["A2"].alignment = Alignment(horizontal="center")

        ws.cell(row=4, column=1, value="TAX INVOICE").font = Font(size=12, bold=True, color="1E3A8A")
        ws.cell(row=5, column=1, value=f"Invoice No: {meta['invoice_number']}").font = bold_font
        ws.cell(row=5, column=5, value=f"Invoice Date: {meta['invoice_date']}").font = bold_font
        ws.cell(row=6, column=1, value=f"Billed To: {meta['client_name']}").font = bold_font
        ws.cell(row=6, column=5, value=f"Client GSTIN: {meta['client_gstin']}").font = bold_font
        ws.cell(row=7, column=1, value=f"Place of Supply: {meta['client_state']}")

        # Table Headers
        headers = ["#", "Item Description", "HSN/SAC", "Qty", "Unit Rate (₹)", "Discount %", "Taxable (₹)", "Total (₹)"]
        start_row = 9
        for col_idx, h in enumerate(headers, start=1):
            c = ws.cell(row=start_row, column=col_idx, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
            c.alignment = Alignment(horizontal="center")

        # Table Rows
        cur_r = start_row + 1
        for item in processed_items:
            ws.cell(row=cur_r, column=1, value=item["index"]).alignment = Alignment(horizontal="center")
            ws.cell(row=cur_r, column=2, value=item["description"])
            ws.cell(row=cur_r, column=3, value=item["hsn"]).alignment = Alignment(horizontal="center")
            ws.cell(row=cur_r, column=4, value=f"{item['qty']} {item['unit']}").alignment = Alignment(horizontal="center")
            ws.cell(row=cur_r, column=5, value=item["rate"]).number_format = "₹#,##0.00"
            ws.cell(row=cur_r, column=6, value=f"{item['discount']}%").alignment = Alignment(horizontal="center")
            ws.cell(row=cur_r, column=7, value=item["taxable"]).number_format = "₹#,##0.00"
            ws.cell(row=cur_r, column=8, value=item["total"]).number_format = "₹#,##0.00"

            for c_i in range(1, 9):
                ws.cell(row=cur_r, column=c_i).border = thin_border
            cur_r += 1

        # Totals
        cur_r += 1
        ws.cell(row=cur_r, column=7, value="Taxable Subtotal:").font = bold_font
        ws.cell(row=cur_r, column=8, value=total_taxable).number_format = "₹#,##0.00"
        cur_r += 1

        if is_interstate:
            ws.cell(row=cur_r, column=7, value="IGST Total:").font = bold_font
            ws.cell(row=cur_r, column=8, value=total_igst).number_format = "₹#,##0.00"
            cur_r += 1
        else:
            ws.cell(row=cur_r, column=7, value="CGST Total:").font = bold_font
            ws.cell(row=cur_r, column=8, value=total_cgst).number_format = "₹#,##0.00"
            cur_r += 1
            ws.cell(row=cur_r, column=7, value="SGST Total:").font = bold_font
            ws.cell(row=cur_r, column=8, value=total_sgst).number_format = "₹#,##0.00"
            cur_r += 1

        ws.cell(row=cur_r, column=7, value="Round Off:").font = bold_font
        ws.cell(row=cur_r, column=8, value=round_off).number_format = "₹#,##0.00"
        cur_r += 1

        ws.cell(row=cur_r, column=7, value="NET PAYABLE:").font = Font(size=12, bold=True, color="1E3A8A")
        ws.cell(row=cur_r, column=8, value=net_payable).font = Font(size=12, bold=True, color="1E3A8A")
        ws.cell(row=cur_r, column=8).number_format = "₹#,##0"

        # Column widths
        widths = [5, 38, 12, 12, 15, 12, 16, 16]
        for col_idx, width in enumerate(widths, start=1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)

        st.download_button(
            label="📊 Download Excel Invoice (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name=f"GST_Invoice_{meta['invoice_number']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with d_col2:
        st.markdown("#### 📄 Printable Vector PDF")
        st.caption("Generate a high-definition PDF ready for direct email delivery to clients or chartered accountants.")
        st.info("💡 **Tip:** Go to the **'👁️ Live Invoice Document Preview'** tab and click the **🖨️ Print / Save as PDF** button. In your browser's print dialog, select **'Save as PDF'** to export an ultra-crisp, watermark-free vector PDF!")

# ==========================================
# FOOTER
# ==========================================
st.markdown(
    """
    <div style="text-align: center; margin-top: 60px; padding: 25px; color: #94a3b8; font-size: 0.85rem; border-top: 1px solid #e2e8f0;">
        <strong>InvoiceCraft GST SaaS Engine &copy; 2026</strong> • Enterprise Billing Automation<br>
        Engineered with Python &amp; Streamlit by <strong>Onkar Hande</strong> (BCA / MCA)<br>
        <a href="https://onkarhande96.github.io" target="_blank" style="color: #2563eb; text-decoration: none; font-weight: 600;">View Portfolio</a> • 
        <a href="https://github.com/Onkarhande96" target="_blank" style="color: #2563eb; text-decoration: none; font-weight: 600;">GitHub</a> • 
        <a href="https://linkedin.com/in/onkarhande96" target="_blank" style="color: #2563eb; text-decoration: none; font-weight: 600;">LinkedIn</a>
    </div>
    """,
    unsafe_allow_html=True,
)
