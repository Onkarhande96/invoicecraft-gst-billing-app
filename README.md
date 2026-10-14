# 🧾 InvoiceCraft — Automated GST Invoice & Billing SaaS

> **A modern, full-featured GST Tax Invoice & Billing Web Application designed for Indian small businesses, freelancers, agencies, and e-commerce traders to generate audit-ready invoices in seconds.**

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://onkar-invoicecraft.streamlit.app)
[![Author](https://img.shields.io/badge/Author-Onkar%20Hande%20(BCA%2FMCA)-blue.svg)](https://onkarhande96.github.io)
[![Python](https://img.shields.io/badge/Python-3.10%2B-green.svg)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-purple.svg)](LICENSE)

---

## 📌 Executive Overview
Small businesses, independent consultants, and retailers often waste hours manually formatting invoices in Word or Excel, risking GST non-compliance, wrong tax calculations, and unprofessional presentations.

**InvoiceCraft** automates this entirely:
1. **Intelligent GST Engine**: Automatically detects whether an invoice is **Intra-State** (splits into CGST + SGST 50/50) or **Inter-State** (computes IGST 100%) by comparing seller and buyer state codes.
2. **Real-Time Taxable Values**: Dynamic line-item addition with HSN/SAC codes, quantity, unit rates, item-level discounts, and multi-tier GST rates (0%, 5%, 12%, 18%, 28%).
3. **Automated Amount in Words**: Built-in Indian numbering currency-to-words generator (e.g. *"Rupees Forty-Seven Thousand Two Hundred Only"*).
4. **Dual Export Formats**:
   - **Printable Vector PDF**: Browser-native 1-click print stylesheet generating high-resolution, watermark-free vector PDFs.
   - **Audit-Ready Excel Workbook (`.xlsx`)**: Auto-formatted workbook using `openpyxl` with financial formulas and cell styling.
5. **Pre-configured Business Presets**: 1-click loading for IT Consulting, E-Commerce Retail, and Inter-State Cloud Services.

---

## 🛠️ Technology Stack
- **Frontend / Application Framework:** [Streamlit](https://streamlit.io/)
- **Data Manipulation:** [Pandas](https://pandas.pydata.org/)
- **Spreadsheet Automation Engine:** [OpenPyXL](https://openpyxl.readthedocs.io/)
- **Styling & Layout:** Responsive CSS3 & Print Media Stylesheets
- **Cloud Deployment:** Streamlit Community Cloud

---

## 🚀 Live Demo & Portfolio
- **Live Cloud Web Application:** [https://onkar-invoicecraft.streamlit.app](https://onkar-invoicecraft.streamlit.app)
- **Developer Portfolio:** [https://onkarhande96.github.io](https://onkarhande96.github.io)

---

## 💻 Local Installation & Setup

```bash
# 1. Clone this repository
git clone https://github.com/Onkarhande96/invoicecraft-gst-billing-app.git

# 2. Navigate to directory
cd invoicecraft-gst-billing-app

# 3. Install required packages
pip install -r requirements.txt

# 4. Launch the application
streamlit run app.py
```

The application will launch on `http://localhost:8501`.

---

## 👨‍💻 Author & Engineering Contact
**Onkar Hande**  
- **Role:** Data Analyst & Workflow Automation Specialist (BCA / MCA)
- **Portfolio:** [https://onkarhande96.github.io](https://onkarhande96.github.io)
- **LinkedIn:** [linkedin.com/in/onkarhande96](https://linkedin.com/in/onkarhande96)
- **Fiverr:** [fiverr.com/onkarhande96](https://fiverr.com/onkarhande96)
- **Email:** onkarhande96@gmail.com
