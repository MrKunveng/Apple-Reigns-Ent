"""
Apple Reigns Enterprise — Billing Suite
Invoice generator · Payment receipts · Revenue analytics.

Design system: Swiss / Minimalism — navy (#1E3A5F) + paid-green (#059669),
Poppins headings / Open Sans body. PDFs rendered with fpdf2 (pure Python,
no system dependencies — deploys cleanly on Streamlit Community Cloud).
"""

from __future__ import annotations

import base64
import os
import re
from datetime import date, datetime, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st
from fpdf import FPDF

# --------------------------------------------------------------------------- #
# Brand / design tokens
# --------------------------------------------------------------------------- #
NAVY = (30, 58, 95)        # #1E3A5F  primary
NAVY_DARK = (36, 58, 73)   # #243A49  logo-background band (blends with logo)
BLUE = (37, 99, 235)       # #2563EB  secondary
GREEN = (5, 150, 105)      # #059669  accent / paid
SLATE = (15, 23, 42)       # #0F172A  foreground
MUTED = (100, 116, 139)    # slate-500
ROW_ALT = (243, 246, 250)  # zebra fill
BORDER = (228, 231, 235)   # #E4E7EB

LOGO_PATH = "assets_logo.png" if os.path.exists("assets_logo.png") else "Logo1.png"
SIGNATURE_PATH = "Signature.png"
TX_FILE = "transactions.csv"
TX_COLS = ["date", "doc_no", "received_from", "payment_for",
           "mode", "reference", "amount", "currency"]

CURRENCIES = ["GHS", "USD", "EUR", "GBP", "NGN"]
PAYMENT_MODES = [
    "Cash",
    "Bank Transfer — CalBank",
    "Bank Transfer — GT Bank",
    "MTN Mobile Money",
    "Cheque",
    "Card / POS",
    "Other",
]

BUSINESS = {
    "name": "Apple Reigns Enterprise",
    "address1": "No. 158 Crane St.",
    "address2": "Accra, Ghana",
    "email": "applereignsenterprise@gmail.com",
    "phone": "+233 24 865 3219",
}

# Bank / mobile-money payment instructions printed on every invoice.
PAYMENT = {
    "banks": [
        {"bank": "CalBank", "name": "Apple Reigns Enterprise",
         "number": "1400008686434", "branch": "Legon"},
        {"bank": "GT Bank", "name": "Apple Reigns Enterprise",
         "number": "3216001003063", "branch": "East Legon"},
    ],
    "momo": {"label": "MTN Mobile Money", "agent": "0596314713",
             "name": "Lukman Kunveng"},
}


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def money(value: float, symbol: str) -> str:
    return f"{symbol} {value:,.2f}"


def _slug(text, fallback="Customer") -> str:
    """Turn a name into a safe file-name fragment."""
    s = re.sub(r"[^\w\s-]", "", str(text)).strip()
    s = re.sub(r"[\s_]+", "-", s)
    return s or fallback


def _s(text) -> str:
    """Make text safe for fpdf2 core (latin-1) fonts."""
    if text is None:
        return ""
    repl = {"—": "-", "–": "-", "’": "'", "‘": "'",
            "“": '"', "”": '"', "…": "...", "•": "-",
            " ": " "}
    text = str(text)
    for k, v in repl.items():
        text = text.replace(k, v)
    return text.encode("latin-1", "replace").decode("latin-1")


def amount_to_words(value: float, currency_name="Ghana Cedis",
                    frac_name="Pesewas") -> str:
    """Render an amount in words (graceful no-op if num2words missing)."""
    try:
        from num2words import num2words
    except Exception:
        return ""
    whole = int(value)
    frac = int(round((value - whole) * 100))
    words = num2words(whole).title() + f" {currency_name}"
    if frac:
        words += " and " + num2words(frac).title() + f" {frac_name}"
    return words + " Only"


def compute_totals(df: pd.DataFrame, tax_rate: float, discount: float):
    df = df.copy()
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce").fillna(0)
    df["Amount"] = df["Quantity"] * df["Unit Price"]
    subtotal = float(df["Amount"].sum())
    taxed_base = max(subtotal - discount, 0)
    tax = taxed_base * tax_rate / 100.0
    total = taxed_base + tax
    return df, subtotal, tax, total


# --------------------------------------------------------------------------- #
# Transaction ledger (revenue inflows)
# Backend: Google Sheets when configured via st.secrets, else local CSV.
# --------------------------------------------------------------------------- #
@st.cache_resource(show_spinner=False)
def _gs_worksheet():
    """Return the 'transactions' worksheet, or None if Sheets isn't set up."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except Exception:
        return None
    try:
        has_secrets = "gcp_service_account" in st.secrets and "gsheets" in st.secrets
    except Exception:
        has_secrets = False  # no secrets.toml present at all
    if not has_secrets:
        return None
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets",
                  "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=scopes)
        client = gspread.authorize(creds)
        ref = st.secrets["gsheets"]["spreadsheet"]
        sh = client.open_by_url(ref) if str(ref).startswith("http") \
            else client.open_by_key(ref)
        try:
            ws = sh.worksheet("transactions")
        except Exception:
            ws = sh.add_worksheet("transactions", rows=2000, cols=len(TX_COLS))
        if ws.row_values(1) != TX_COLS:
            if not ws.row_values(1):
                ws.append_row(TX_COLS)
        return ws
    except Exception as exc:  # bad creds / not shared / API off
        st.session_state["_gs_error"] = str(exc)
        return None


def using_gsheets() -> bool:
    return _gs_worksheet() is not None


def backend_label() -> str:
    return "Google Sheets (cloud)" if using_gsheets() else "local CSV file"


@st.cache_data(ttl=20, show_spinner=False)
def _read_gsheet():
    ws = _gs_worksheet()
    if ws is None:
        return None
    return ws.get_all_records(expected_headers=TX_COLS)


def load_transactions() -> pd.DataFrame:
    records = _read_gsheet()
    if records is not None:
        df = pd.DataFrame(records)
        for c in TX_COLS:
            if c not in df.columns:
                df[c] = None
        return df[TX_COLS] if not df.empty else pd.DataFrame(columns=TX_COLS)
    if os.path.exists(TX_FILE):
        try:
            df = pd.read_csv(TX_FILE)
        except Exception:
            return pd.DataFrame(columns=TX_COLS)
        for c in TX_COLS:
            if c not in df.columns:
                df[c] = None
        return df[TX_COLS]
    return pd.DataFrame(columns=TX_COLS)


def record_transaction(row: dict) -> None:
    ws = _gs_worksheet()
    if ws is not None:
        ws.append_row([row.get(c, "") for c in TX_COLS],
                      value_input_option="USER_ENTERED")
        _read_gsheet.clear()
        return
    df = pd.DataFrame(columns=TX_COLS)
    if os.path.exists(TX_FILE):
        try:
            df = pd.read_csv(TX_FILE)
        except Exception:
            pass
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_csv(TX_FILE, index=False)


# --------------------------------------------------------------------------- #
# Shared PDF header
# --------------------------------------------------------------------------- #
class DocPDF(FPDF):
    def __init__(self):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.set_auto_page_break(auto=True, margin=18)
        self.set_margins(15, 15, 15)

    def footer(self):
        self.set_y(-15)
        self.set_draw_color(*BORDER)
        self.set_line_width(0.2)
        self.line(15, self.get_y(), 195, self.get_y())
        self.set_y(-12)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(*MUTED)
        contact = f"{BUSINESS['name']}  |  {BUSINESS['email']}  |  {BUSINESS['phone']}"
        self.cell(0, 6, contact, align="C")


def _draw_header(pdf: FPDF, title: str, meta: list[tuple[str, str]]) -> float:
    """Navy band + logo + right-aligned title and meta. Returns band height."""
    page_w, band_h = 210, 46
    pdf.set_fill_color(*NAVY_DARK)
    pdf.rect(0, 0, page_w, band_h, style="F")
    pdf.set_fill_color(*GREEN)
    pdf.rect(0, band_h, page_w, 1.4, style="F")
    try:
        pdf.image(LOGO_PATH, x=15, y=7, h=32)
    except Exception:
        pass
    pdf.set_xy(115, 11)
    pdf.set_font("Helvetica", "B", 30)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(80, 12, title, align="R")
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(205, 215, 225)
    y = 26
    for label, value in meta:
        pdf.set_xy(115, y)
        pdf.cell(80, 5.5, f"{label}  {value}", align="R")
        y += 5.5
    return band_h


def _secret_get(section: str, key: str, default=None):
    """Read st.secrets[section][key] without crashing when no secrets exist."""
    try:
        if section in st.secrets and key in st.secrets[section]:
            return st.secrets[section][key]
    except Exception:
        pass
    return default


def signature_b64() -> str:
    """Signature image as base64 — from secrets first, else the local file."""
    data = _secret_get("signature", "data")
    if data:
        return str(data).strip()
    if os.path.exists(SIGNATURE_PATH):
        with open(SIGNATURE_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""


def signature_bytes():
    b64 = signature_b64()
    if not b64:
        return None
    try:
        return base64.b64decode(b64)
    except Exception:
        return None


def _draw_signature(pdf: FPDF, right: float = 195) -> None:
    """Signature image + line + caption, right-aligned at the current y."""
    sig_w = 70
    sig_x = right - sig_w
    y = pdf.get_y()
    sig = signature_bytes()
    if sig:
        try:
            img_w = 42
            pdf.image(BytesIO(sig), x=sig_x + (sig_w - img_w) / 2, y=y, w=img_w)
        except Exception:
            pass
    pdf.set_y(y + 18)
    pdf.set_draw_color(*MUTED)
    pdf.set_line_width(0.3)
    pdf.line(sig_x, pdf.get_y(), right, pdf.get_y())
    pdf.set_xy(sig_x, pdf.get_y() + 1.5)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(*MUTED)
    pdf.multi_cell(sig_w, 5, "Authorised Signature\nFor: " + BUSINESS["name"],
                   align="C")


# --------------------------------------------------------------------------- #
# Invoice PDF
# --------------------------------------------------------------------------- #
def build_pdf(df, *, invoice_no, invoice_date, bill_to_name, bill_to_lines,
              symbol, subtotal, tax, tax_rate, discount, total, notes,
              payment_terms) -> bytes:
    invoice_no, invoice_date, bill_to_name = _s(invoice_no), _s(invoice_date), _s(bill_to_name)
    bill_to_lines = [_s(l) for l in bill_to_lines]
    notes, payment_terms = _s(notes), _s(payment_terms)
    df = df.copy()
    df["Description"] = df["Description"].map(_s)

    pdf = DocPDF()
    pdf.add_page()
    left, right = 15, 195
    inner = right - left
    band_h = _draw_header(pdf, "INVOICE",
                          [("Invoice No.", invoice_no), ("Date", invoice_date)])

    # ---- From / Bill To --------------------------------------------------- #
    pdf.set_y(band_h + 9)
    col_w = inner / 2
    top = pdf.get_y()
    pdf.set_xy(left, top)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*GREEN)
    pdf.cell(col_w, 5, "FROM")
    pdf.set_xy(left, top + 6)
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(*NAVY)
    pdf.cell(col_w, 6, BUSINESS["name"])
    pdf.set_font("Helvetica", "", 9.5)
    pdf.set_text_color(*MUTED)
    for i, line in enumerate([BUSINESS["address1"], BUSINESS["address2"],
                              BUSINESS["email"], BUSINESS["phone"]]):
        pdf.set_xy(left, top + 13 + i * 5)
        pdf.cell(col_w, 5, line)

    bx = left + col_w
    pdf.set_xy(bx, top)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*GREEN)
    pdf.cell(col_w, 5, "BILL TO")
    pdf.set_xy(bx, top + 6)
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(*NAVY)
    pdf.cell(col_w, 6, bill_to_name)
    pdf.set_font("Helvetica", "", 9.5)
    pdf.set_text_color(*MUTED)
    for i, line in enumerate([l for l in bill_to_lines if l.strip()]):
        pdf.set_xy(bx, top + 13 + i * 5)
        pdf.cell(col_w, 5, line)

    # ---- Items table ------------------------------------------------------ #
    pdf.set_y(top + 40)
    w_desc = inner * 0.50
    w_qty = inner * 0.12
    w_price = inner * 0.19
    w_amt = inner * 0.19
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(w_desc, 9, "  DESCRIPTION", align="L", fill=True)
    pdf.cell(w_qty, 9, "QTY", align="C", fill=True)
    pdf.cell(w_price, 9, "UNIT PRICE", align="R", fill=True)
    pdf.cell(w_amt, 9, "AMOUNT  ", align="R", fill=True, ln=1)

    pdf.set_font("Helvetica", "", 9.5)
    for idx, (_, row) in enumerate(df.iterrows()):
        desc = str(row["Description"]).strip()
        if not desc:
            continue
        fill = idx % 2 == 1
        if fill:
            pdf.set_fill_color(*ROW_ALT)
        pdf.set_text_color(*SLATE)
        pdf.cell(w_desc, 8, "  " + desc, align="L", fill=fill)
        pdf.cell(w_qty, 8, f"{float(row['Quantity']):g}", align="C", fill=fill)
        pdf.cell(w_price, 8, money(float(row["Unit Price"]), symbol),
                 align="R", fill=fill)
        pdf.cell(w_amt, 8, money(float(row["Amount"]), symbol) + "  ",
                 align="R", fill=fill, ln=1)

    pdf.set_draw_color(*NAVY)
    pdf.set_line_width(0.4)
    y = pdf.get_y()
    pdf.line(left, y, right, y)

    # ---- Totals ----------------------------------------------------------- #
    pdf.ln(4)
    label_w = inner * 0.62
    val_w = inner * 0.38

    def total_row(label, value, *, bold=False, accent=False, size=10):
        pdf.set_x(left)
        pdf.set_font("Helvetica", "B" if bold else "", size)
        pdf.set_text_color(*MUTED)
        pdf.cell(label_w, 7, "")
        if accent:
            pdf.set_fill_color(*GREEN)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(val_w * 0.45, 9, "  " + label, align="L", fill=True)
            pdf.cell(val_w * 0.55, 9, value + "  ", align="R", fill=True, ln=1)
        else:
            pdf.cell(val_w * 0.45, 7, label, align="L")
            pdf.set_text_color(*SLATE)
            pdf.cell(val_w * 0.55, 7, value, align="R", ln=1)

    total_row("Subtotal", money(subtotal, symbol))
    if discount > 0:
        total_row("Discount", "-" + money(discount, symbol))
    total_row(f"VAT ({tax_rate:g}%)", money(tax, symbol))
    pdf.ln(1)
    total_row("TOTAL", money(total, symbol), bold=True, accent=True, size=12)

    # ---- Payment details panel ------------------------------------------- #
    pdf.ln(9)
    box_y = pdf.get_y()
    box_h = 47
    pdf.set_fill_color(248, 250, 252)
    pdf.set_draw_color(*BORDER)
    pdf.set_line_width(0.3)
    pdf.rect(left, box_y, inner, box_h, style="DF")
    pdf.set_fill_color(*GREEN)
    pdf.rect(left, box_y, 1.6, box_h, style="F")

    pad = 7
    px, py = left + pad, box_y + 6
    pdf.set_xy(px, py)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*GREEN)
    pdf.cell(inner - pad, 5, "PAYMENT DETAILS")

    bank_col = (inner - pad * 2) / 2
    by = py + 8
    for i, bank in enumerate(PAYMENT["banks"]):
        cx = px + i * bank_col
        pdf.set_xy(cx, by)
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(*NAVY)
        pdf.cell(bank_col, 5, bank["bank"])
        pdf.set_font("Helvetica", "", 8.6)
        pdf.set_text_color(*MUTED)
        for j, d in enumerate([f"Account Name: {bank['name']}",
                               f"Account Number: {bank['number']}",
                               f"Branch: {bank['branch']}"]):
            pdf.set_xy(cx, by + 5.5 + j * 4.6)
            pdf.cell(bank_col, 4.4, d)

    my = by + 26
    pdf.set_draw_color(*BORDER)
    pdf.line(px, my - 2, left + inner - pad, my - 2)
    momo = PAYMENT["momo"]
    pdf.set_xy(px, my)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(*NAVY)
    pdf.cell(bank_col, 5, momo["label"])
    pdf.set_font("Helvetica", "", 8.6)
    pdf.set_text_color(*MUTED)
    pdf.set_xy(px, my + 5.5)
    pdf.cell(bank_col * 2, 4.4,
             f"Agent Number: {momo['agent']}    Name: {momo['name']}")
    pdf.set_y(box_y + box_h)

    # ---- Notes / terms ---------------------------------------------------- #
    pdf.ln(8)
    if payment_terms.strip():
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_text_color(*NAVY)
        pdf.set_x(left)
        pdf.cell(inner, 5, "Payment Terms", ln=1)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(*MUTED)
        pdf.set_x(left)
        pdf.multi_cell(inner, 5, payment_terms)
        pdf.ln(2)
    if notes.strip():
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_text_color(*NAVY)
        pdf.set_x(left)
        pdf.cell(inner, 5, "Notes", ln=1)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(*MUTED)
        pdf.set_x(left)
        pdf.multi_cell(inner, 5, notes)

    # signature (right-aligned)
    pdf.ln(8)
    _draw_signature(pdf, right)

    pdf.ln(6)
    pdf.set_font("Helvetica", "I", 11)
    pdf.set_text_color(*GREEN)
    pdf.set_x(left)
    pdf.cell(inner, 7, "Thank you for your business.", align="C")
    return bytes(pdf.output())


# --------------------------------------------------------------------------- #
# Receipt PDF
# --------------------------------------------------------------------------- #
def build_receipt_pdf(*, receipt_no, receipt_date, received_from, payment_for,
                      mode, reference, amount, symbol, words) -> bytes:
    receipt_no, receipt_date = _s(receipt_no), _s(receipt_date)
    received_from, payment_for = _s(received_from), _s(payment_for)
    mode, reference, words = _s(mode), _s(reference), _s(words)

    pdf = DocPDF()
    pdf.add_page()
    left, right = 15, 195
    inner = right - left
    band_h = _draw_header(pdf, "RECEIPT",
                          [("Receipt No.", receipt_no), ("Date", receipt_date)])

    pdf.set_y(band_h + 14)
    pdf.set_x(left)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*GREEN)
    pdf.cell(inner, 5, "RECEIVED WITH THANKS FROM", ln=1)
    pdf.set_x(left)
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_text_color(*NAVY)
    pdf.cell(inner, 9, received_from, ln=1)
    pdf.ln(4)

    # amount badge
    badge_y = pdf.get_y()
    pdf.set_fill_color(*GREEN)
    pdf.rect(left, badge_y, inner, 22, style="F")
    pdf.set_xy(left + 7, badge_y + 4)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(230, 245, 238)
    pdf.cell(100, 5, "AMOUNT RECEIVED")
    pdf.set_xy(left + 7, badge_y + 9.5)
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(120, 10, money(amount, symbol))
    # PAID stamp
    pdf.set_xy(right - 47, badge_y + 6)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(255, 255, 255)
    pdf.set_draw_color(255, 255, 255)
    pdf.set_line_width(0.6)
    pdf.cell(40, 10, "PAID", border=1, align="C")
    pdf.set_y(badge_y + 26)

    if words:
        pdf.set_x(left)
        pdf.set_font("Helvetica", "I", 9.5)
        pdf.set_text_color(*MUTED)
        pdf.multi_cell(inner, 5, "Amount in words: " + words)
        pdf.ln(2)

    def detail(label, value):
        pdf.set_x(left)
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_text_color(*MUTED)
        pdf.cell(48, 8, label)
        pdf.set_font("Helvetica", "", 10.5)
        pdf.set_text_color(*SLATE)
        pdf.multi_cell(inner - 48, 8, value if str(value).strip() else "-")

    pdf.ln(2)
    detail("Being payment for:", payment_for)
    detail("Mode of payment:", mode)
    detail("Reference / Txn ID:", reference)
    detail("Date:", receipt_date)

    # signature
    pdf.ln(16)
    _draw_signature(pdf, right)
    return bytes(pdf.output())


# --------------------------------------------------------------------------- #
# HTML previews
# --------------------------------------------------------------------------- #
def _responsive_doc(inner: str) -> str:
    """Wrap preview HTML in a mobile-aware document (viewport + fluid padding)."""
    return (
        "<!doctype html><html><head>"
        "<meta name='viewport' content='width=device-width, initial-scale=1'>"
        "<style>*{box-sizing:border-box}html,body{margin:0}"
        "body{background:#EEF2F6;padding:16px;"
        "font-family:'Open Sans',system-ui,sans-serif}"
        "@media(max-width:480px){body{padding:8px}}</style></head>"
        f"<body>{inner}</body></html>"
    )


def _sig_html() -> str:
    return f"""
      <div style="display:flex;justify-content:flex-end;padding:16px clamp(16px,5vw,32px) 6px">
        <div style="text-align:center;color:#64748B;font-size:13px">
          <img src="data:image/png;base64,{st.session_state.get('sig_b64', '')}"
               style="height:46px;margin-bottom:2px" alt="signature"/>
          <div style="width:200px;border-top:1px solid #94A3B8;padding-top:6px">Authorised Signature</div>
          <div style="margin-top:2px">For: {BUSINESS['name']}</div>
        </div>
      </div>"""


def render_invoice_preview(df, *, invoice_no, invoice_date, bill_to_name,
                           bill_to_lines, symbol, subtotal, tax, tax_rate,
                           discount, total, notes, payment_terms):
    rows = ""
    for idx, (_, r) in enumerate(df.iterrows()):
        desc = str(r["Description"]).strip()
        if not desc:
            continue
        bg = "#F3F6FA" if idx % 2 else "#FFFFFF"
        rows += f"""
        <tr style="background:{bg}">
          <td style="padding:11px 14px;color:#0F172A">{desc}</td>
          <td style="padding:11px 14px;text-align:center;color:#0F172A">{float(r['Quantity']):g}</td>
          <td style="padding:11px 14px;text-align:right;color:#0F172A">{money(float(r['Unit Price']), symbol)}</td>
          <td style="padding:11px 14px;text-align:right;color:#0F172A;font-weight:600">{money(float(r['Amount']), symbol)}</td>
        </tr>"""

    billto = "<br>".join(l for l in bill_to_lines if l.strip())
    bank_cards = ""
    for b in PAYMENT["banks"]:
        bank_cards += f"""
          <div style="flex:1;min-width:170px">
            <div style="font-family:'Poppins',sans-serif;color:#1E3A5F;font-weight:600;font-size:14px">{b['bank']}</div>
            <div style="color:#64748B;font-size:12.5px;line-height:1.7;margin-top:3px">
              Account Name: {b['name']}<br>Account Number: {b['number']}<br>Branch: {b['branch']}
            </div>
          </div>"""
    momo = PAYMENT["momo"]
    payment_panel = f"""
      <div style="margin:0 clamp(16px,5vw,32px) 18px;background:#F8FAFC;border:1px solid #E4E7EB;
                  border-left:4px solid #059669;border-radius:10px;padding:16px 20px">
        <div style="color:#059669;font-weight:700;font-size:12px;letter-spacing:.8px;margin-bottom:10px">PAYMENT DETAILS</div>
        <div style="display:flex;flex-wrap:wrap;gap:24px">{bank_cards}</div>
        <div style="border-top:1px solid #E4E7EB;margin-top:12px;padding-top:10px">
          <span style="font-family:'Poppins',sans-serif;color:#1E3A5F;font-weight:600;font-size:14px">{momo['label']}</span>
          <span style="color:#64748B;font-size:12.5px;margin-left:10px">Agent Number: {momo['agent']} &nbsp;&middot;&nbsp; Name: {momo['name']}</span>
        </div>
      </div>"""

    discount_row = (
        f'<tr><td style="padding:4px 0;color:#64748B">Discount</td>'
        f'<td style="padding:4px 0;text-align:right;color:#0F172A">-{money(discount, symbol)}</td></tr>'
        if discount > 0 else ""
    )
    terms_html = (f"<div style='padding:0 clamp(16px,5vw,32px) 10px'><div style='color:#1E3A5F;font-weight:600;font-size:13px'>Payment Terms</div>"
                  f"<div style='color:#64748B;font-size:13px'>{payment_terms}</div></div>" if payment_terms.strip() else "")
    notes_html = (f"<div style='padding:0 clamp(16px,5vw,32px) 10px'><div style='color:#1E3A5F;font-weight:600;font-size:13px'>Notes</div>"
                  f"<div style='color:#64748B;font-size:13px'>{notes}</div></div>" if notes.strip() else "")

    return f"""
    <div style="font-family:'Open Sans',system-ui,sans-serif;max-width:780px;margin:0 auto;
                width:100%;background:#fff;border:1px solid #E4E7EB;border-radius:14px;overflow:hidden;
                box-shadow:0 10px 30px rgba(15,23,42,.08)">
      <div style="background:#243A49;padding:26px clamp(16px,5vw,32px);display:flex;justify-content:space-between;
                  flex-wrap:wrap;gap:12px;align-items:center;border-bottom:3px solid #059669">
        <img src="data:image/png;base64,{st.session_state['logo_b64']}" style="height:clamp(50px,14vw,74px)" alt="logo"/>
        <div style="text-align:right;color:#fff">
          <div style="font-family:'Poppins',sans-serif;font-size:clamp(24px,7vw,34px);font-weight:700;letter-spacing:1px">INVOICE</div>
          <div style="color:#CDD7E1;font-size:14px;margin-top:6px">Invoice No.&nbsp;&nbsp;{invoice_no}</div>
          <div style="color:#CDD7E1;font-size:14px">Date&nbsp;&nbsp;{invoice_date}</div>
        </div>
      </div>
      <div style="display:flex;flex-wrap:wrap;gap:24px;padding:28px clamp(16px,5vw,32px) 6px">
        <div style="flex:1;min-width:170px">
          <div style="color:#059669;font-weight:700;font-size:12px;letter-spacing:.8px">FROM</div>
          <div style="font-family:'Poppins',sans-serif;color:#1E3A5F;font-weight:600;font-size:16px;margin:5px 0 6px">{BUSINESS['name']}</div>
          <div style="color:#64748B;font-size:13px;line-height:1.7">{BUSINESS['address1']}<br>{BUSINESS['address2']}<br>{BUSINESS['email']}<br>{BUSINESS['phone']}</div>
        </div>
        <div style="flex:1;min-width:170px">
          <div style="color:#059669;font-weight:700;font-size:12px;letter-spacing:.8px">BILL TO</div>
          <div style="font-family:'Poppins',sans-serif;color:#1E3A5F;font-weight:600;font-size:16px;margin:5px 0 6px">{bill_to_name}</div>
          <div style="color:#64748B;font-size:13px;line-height:1.7">{billto}</div>
        </div>
      </div>
      <div style="padding:18px clamp(16px,5vw,32px) 8px;overflow-x:auto">
        <table style="width:100%;border-collapse:collapse;font-size:13.5px;min-width:420px">
          <thead><tr style="background:#1E3A5F;color:#fff;font-family:'Poppins',sans-serif">
            <th style="padding:12px 14px;text-align:left;font-weight:600">DESCRIPTION</th>
            <th style="padding:12px 14px;text-align:center;font-weight:600">QTY</th>
            <th style="padding:12px 14px;text-align:right;font-weight:600">UNIT PRICE</th>
            <th style="padding:12px 14px;text-align:right;font-weight:600">AMOUNT</th>
          </tr></thead>
          <tbody>{rows}</tbody>
        </table>
      </div>
      <div style="display:flex;justify-content:flex-end;padding:6px clamp(16px,5vw,32px) 22px">
        <table style="font-size:14px;min-width:280px">
          <tr><td style="padding:4px 0;color:#64748B">Subtotal</td><td style="padding:4px 0;text-align:right;color:#0F172A">{money(subtotal, symbol)}</td></tr>
          {discount_row}
          <tr><td style="padding:4px 0;color:#64748B">VAT ({tax_rate:g}%)</td><td style="padding:4px 0;text-align:right;color:#0F172A">{money(tax, symbol)}</td></tr>
          <tr><td colspan="2" style="padding-top:8px">
            <div style="display:flex;justify-content:space-between;background:#059669;color:#fff;padding:12px 16px;border-radius:8px;font-family:'Poppins',sans-serif;font-weight:700;font-size:16px">
              <span>TOTAL</span><span>{money(total, symbol)}</span></div>
          </td></tr>
        </table>
      </div>
      {payment_panel}
      {terms_html}
      {notes_html}
      {_sig_html()}
      <div style="background:#F8FAFC;border-top:1px solid #E4E7EB;padding:16px clamp(16px,5vw,32px);text-align:center;color:#059669;font-style:italic;font-size:14px">Thank you for your business.</div>
    </div>
    """


def render_receipt_preview(*, receipt_no, receipt_date, received_from,
                           payment_for, mode, reference, amount, symbol, words):
    def detail(label, value):
        return f"""<tr>
            <td style="padding:9px 0;color:#64748B;font-weight:600;width:45%;vertical-align:top">{label}</td>
            <td style="padding:9px 0;color:#0F172A">{value if str(value).strip() else '-'}</td></tr>"""

    words_html = (f"<div style='color:#64748B;font-style:italic;font-size:13px;margin:14px 0 4px'>Amount in words: {words}</div>"
                  if words else "")

    return f"""
    <div style="font-family:'Open Sans',system-ui,sans-serif;max-width:780px;margin:0 auto;
                width:100%;background:#fff;border:1px solid #E4E7EB;border-radius:14px;overflow:hidden;
                box-shadow:0 10px 30px rgba(15,23,42,.08)">
      <div style="background:#243A49;padding:26px clamp(16px,5vw,32px);display:flex;justify-content:space-between;
                  flex-wrap:wrap;gap:12px;align-items:center;border-bottom:3px solid #059669">
        <img src="data:image/png;base64,{st.session_state['logo_b64']}" style="height:clamp(50px,14vw,74px)" alt="logo"/>
        <div style="text-align:right;color:#fff">
          <div style="font-family:'Poppins',sans-serif;font-size:clamp(24px,7vw,34px);font-weight:700;letter-spacing:1px">RECEIPT</div>
          <div style="color:#CDD7E1;font-size:14px;margin-top:6px">Receipt No.&nbsp;&nbsp;{receipt_no}</div>
          <div style="color:#CDD7E1;font-size:14px">Date&nbsp;&nbsp;{receipt_date}</div>
        </div>
      </div>
      <div style="padding:26px clamp(16px,5vw,32px) 4px">
        <div style="color:#059669;font-weight:700;font-size:12px;letter-spacing:.8px">RECEIVED WITH THANKS FROM</div>
        <div style="font-family:'Poppins',sans-serif;color:#1E3A5F;font-weight:600;font-size:22px;margin-top:4px">{received_from}</div>
      </div>
      <div style="margin:18px clamp(16px,5vw,32px);background:#059669;border-radius:12px;padding:18px 22px;
                  display:flex;flex-wrap:wrap;gap:12px;justify-content:space-between;align-items:center;color:#fff">
        <div>
          <div style="font-size:12px;letter-spacing:.8px;opacity:.85">AMOUNT RECEIVED</div>
          <div style="font-family:'Poppins',sans-serif;font-size:clamp(22px,6vw,30px);font-weight:700;margin-top:2px">{money(amount, symbol)}</div>
        </div>
        <div style="border:2px solid #fff;border-radius:8px;padding:6px 18px;font-family:'Poppins',sans-serif;font-weight:700;font-size:18px;letter-spacing:1px">PAID</div>
      </div>
      <div style="padding:0 clamp(16px,5vw,32px) 8px">{words_html}
        <table style="width:100%;border-collapse:collapse;font-size:14px;margin-top:6px">
          {detail("Being payment for:", payment_for)}
          {detail("Mode of payment:", mode)}
          {detail("Reference / Txn ID:", reference)}
          {detail("Date:", receipt_date)}
        </table>
      </div>
      {_sig_html()}
      <div style="background:#F8FAFC;border-top:1px solid #E4E7EB;padding:16px clamp(16px,5vw,32px);text-align:center;color:#059669;font-style:italic;font-size:14px">Thank you for your payment.</div>
    </div>
    """


# --------------------------------------------------------------------------- #
# Views
# --------------------------------------------------------------------------- #
def invoice_view():
    with st.sidebar:
        st.header("Invoice details")
        invoice_no = st.text_input("Invoice number", value="005")
        invoice_date = st.date_input("Invoice date", value=date(2026, 6, 19),
                                     key="inv_date")
        date_str = invoice_date.strftime("%B %d, %Y")
        st.subheader("Bill to")
        bill_to_name = st.text_input("Client / organisation",
                                     value="Lawra Municipal Assembly")
        bill_to_extra = st.text_area("Client address (one line each)", value="",
                                     placeholder="Street\nCity, Region\nemail / phone",
                                     height=90)
        st.subheader("Settings")
        symbol = st.selectbox("Currency", CURRENCIES, index=0, key="inv_ccy")
        tax_rate = st.number_input("VAT / Tax (%)", 0.0, 100.0, 0.0, 0.5)
        discount = st.number_input("Discount", 0.0, value=0.0, step=10.0)
        st.subheader("Footer")
        payment_terms = st.text_area(
            "Payment terms",
            value="Payment due within 30 days. Bank/MoMo details above.",
            height=80)
        notes = st.text_area("Notes", value="", height=70)

    st.subheader("Line items")
    default_items = pd.DataFrame(
        [{"Description": "Service / product description",
          "Quantity": 1, "Unit Price": 0.00}])
    edited = st.data_editor(
        default_items, num_rows="dynamic", use_container_width=True, key="items",
        column_config={
            "Description": st.column_config.TextColumn("Description", width="large"),
            "Quantity": st.column_config.NumberColumn("Qty", min_value=0, step=1, format="%g"),
            "Unit Price": st.column_config.NumberColumn("Unit Price", min_value=0.0, step=10.0, format="%.2f"),
        })

    df, subtotal, tax, total = compute_totals(edited, tax_rate, discount)
    bill_to_lines = bill_to_extra.splitlines()

    c1, c2, c3 = st.columns(3)
    c1.metric("Subtotal", money(subtotal, symbol))
    c2.metric(f"VAT ({tax_rate:g}%)", money(tax, symbol))
    c3.metric("Total", money(total, symbol))
    st.divider()

    left, right = st.columns([2, 1])
    with right:
        st.subheader("Export")
        st.caption("Generate a print-ready A4 PDF.")
        if df["Description"].astype(str).str.strip().ne("").any():
            pdf_bytes = build_pdf(
                df, invoice_no=invoice_no, invoice_date=date_str,
                bill_to_name=bill_to_name, bill_to_lines=bill_to_lines,
                symbol=symbol, subtotal=subtotal, tax=tax, tax_rate=tax_rate,
                discount=discount, total=total, notes=notes,
                payment_terms=payment_terms)
            st.download_button(
                "Download invoice PDF", data=pdf_bytes,
                file_name=f"{_slug(bill_to_name)}-Invoice-{invoice_no}.pdf",
                mime="application/pdf", use_container_width=True)
        else:
            st.info("Add at least one line item to enable the PDF download.")

    with left:
        st.subheader("Live preview")
        st.components.v1.html(
            _responsive_doc(render_invoice_preview(
                df, invoice_no=invoice_no, invoice_date=date_str,
                bill_to_name=bill_to_name, bill_to_lines=bill_to_lines,
                symbol=symbol, subtotal=subtotal, tax=tax, tax_rate=tax_rate,
                discount=discount, total=total, notes=notes,
                payment_terms=payment_terms)),
            height=1050, scrolling=True)


def receipt_view():
    txns = load_transactions()
    next_no = f"RCP-{len(txns) + 1:03d}"

    with st.sidebar:
        st.header("Receipt details")
        receipt_no = st.text_input("Receipt number", value=next_no)
        receipt_date = st.date_input("Payment date", value=date.today(),
                                     key="rcp_date")
        date_str = receipt_date.strftime("%B %d, %Y")
        received_from = st.text_input("Received from",
                                      value="Lawra Municipal Assembly")
        amount = st.number_input("Amount received", min_value=0.0, value=0.0,
                                 step=50.0, format="%.2f")
        symbol = st.selectbox("Currency", CURRENCIES, index=0, key="rcp_ccy")
        mode = st.selectbox("Mode of payment", PAYMENT_MODES, index=0)
        payment_for = st.text_input("Being payment for",
                                    placeholder="e.g. Invoice #005 — supplies")
        reference = st.text_input("Reference / Txn ID",
                                  placeholder="cheque no, MoMo ref, etc.")

    words = amount_to_words(
        amount, currency_name=("Ghana Cedis" if symbol == "GHS" else symbol),
        frac_name=("Pesewas" if symbol == "GHS" else "Cents")) if amount else ""

    left, right = st.columns([2, 1])
    with right:
        st.subheader("Export & record")
        ready = amount > 0 and received_from.strip()
        if ready:
            pdf_bytes = build_receipt_pdf(
                receipt_no=receipt_no, receipt_date=date_str,
                received_from=received_from, payment_for=payment_for,
                mode=mode, reference=reference, amount=amount, symbol=symbol,
                words=words)
            st.download_button("Download receipt PDF", data=pdf_bytes,
                               file_name=f"Apple-Reigns-Receipt-{receipt_no}.pdf",
                               mime="application/pdf", use_container_width=True)
            st.caption("Logging adds this payment to your revenue analytics.")
            if st.button("Record this payment", type="primary",
                         use_container_width=True):
                record_transaction({
                    "date": receipt_date.isoformat(), "doc_no": receipt_no,
                    "received_from": received_from, "payment_for": payment_for,
                    "mode": mode, "reference": reference, "amount": amount,
                    "currency": symbol})
                st.success(f"Recorded {money(amount, symbol)} from "
                           f"{received_from} → {backend_label()}.")
        else:
            st.info("Enter a payer and an amount above to enable the receipt.")

    with left:
        st.subheader("Live preview")
        st.components.v1.html(
            _responsive_doc(render_receipt_preview(
                receipt_no=receipt_no, receipt_date=date_str,
                received_from=received_from or "—", payment_for=payment_for,
                mode=mode, reference=reference, amount=amount, symbol=symbol,
                words=words)),
            height=880, scrolling=True)


def analytics_view():
    st.subheader("Revenue analytics")
    st.caption(f"Ledger storage: **{backend_label()}**.")
    df = load_transactions()
    if df.empty:
        st.info("No payments recorded yet. Record payments from the **Receipt** "
                "tab (or log one below) to start tracking revenue inflows.")
    else:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
        df = df.dropna(subset=["date"])
        symbol = df["currency"].mode().iat[0] if not df["currency"].isna().all() else "GHS"

        today = pd.Timestamp(date.today())
        last7 = df[df["date"] >= today - pd.Timedelta(days=6)]
        last30 = df[df["date"] >= today - pd.Timedelta(days=29)]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Last 7 days", money(last7["amount"].sum(), symbol),
                  f"{len(last7)} payments")
        m2.metric("Last 30 days", money(last30["amount"].sum(), symbol),
                  f"{len(last30)} payments")
        m3.metric("All-time revenue", money(df["amount"].sum(), symbol),
                  f"{len(df)} payments")
        avg = df["amount"].mean() if len(df) else 0
        m4.metric("Average payment", money(avg, symbol))

        st.divider()
        ca, cb = st.columns(2)
        with ca:
            st.markdown("**Daily revenue — last 30 days**")
            idx = pd.date_range(today - pd.Timedelta(days=29), today)
            daily = (last30.groupby(last30["date"].dt.normalize())["amount"]
                     .sum().reindex(idx, fill_value=0))
            daily.index = daily.index.strftime("%b %d")
            st.bar_chart(daily, color=(5, 150, 105), height=260)
        with cb:
            st.markdown("**Revenue by payment mode — last 30 days**")
            by_mode = (last30.groupby("mode")["amount"].sum()
                       .sort_values(ascending=False))
            if by_mode.empty:
                st.caption("No payments in the last 30 days.")
            else:
                st.bar_chart(by_mode, color=(30, 58, 95), height=260)

        st.markdown("**Recent payments**")
        recent = df.sort_values("date", ascending=False).head(25).copy()
        recent["date"] = recent["date"].dt.strftime("%Y-%m-%d")
        recent = recent.rename(columns={
            "date": "Date", "doc_no": "Receipt", "received_from": "From",
            "payment_for": "For", "mode": "Mode", "reference": "Ref",
            "amount": "Amount", "currency": "Ccy"})
        st.dataframe(recent, use_container_width=True, hide_index=True)
        st.download_button("Download ledger (CSV)",
                           data=df.to_csv(index=False).encode(),
                           file_name="apple-reigns-transactions.csv",
                           mime="text/csv")

    # manual entry — for payments received without a generated receipt
    with st.expander("➕ Log a payment manually"):
        with st.form("manual_tx", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            d = c1.date_input("Date", value=date.today(), key="man_date")
            who = c2.text_input("Received from")
            amt = c3.number_input("Amount", min_value=0.0, value=0.0, step=50.0)
            c4, c5, c6 = st.columns(3)
            ccy = c4.selectbox("Currency", CURRENCIES, index=0, key="man_ccy")
            md = c5.selectbox("Mode", PAYMENT_MODES, index=0, key="man_mode")
            ref = c6.text_input("Reference")
            what = st.text_input("Payment for")
            if st.form_submit_button("Add to ledger", type="primary"):
                if amt > 0 and who.strip():
                    record_transaction({
                        "date": d.isoformat(), "doc_no": "manual",
                        "received_from": who, "payment_for": what, "mode": md,
                        "reference": ref, "amount": amt, "currency": ccy})
                    st.success(f"Logged {money(amt, ccy)} from {who}.")
                    st.rerun()
                else:
                    st.warning("Enter a payer and an amount greater than zero.")


# --------------------------------------------------------------------------- #
# App
# --------------------------------------------------------------------------- #
def main():
    st.set_page_config(page_title="Apple Reigns — Billing Suite",
                       page_icon="🧾", layout="wide")

    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;500;600;700&family=Poppins:wght@400;500;600;700&display=swap');
        html, body, [class*="css"] { font-family:'Open Sans',sans-serif; }
        h1,h2,h3 { font-family:'Poppins',sans-serif !important; color:#1E3A5F; }
        .stButton>button, .stDownloadButton>button, .stFormSubmitButton>button {
            background:#1E3A5F;color:#fff;border:0;border-radius:8px;font-weight:600;
            padding:.55rem 1.1rem;transition:all .2s ease; }
        .stDownloadButton>button { background:#059669; }
        .stButton>button:hover { background:#2563EB; }
        .stDownloadButton>button:hover { background:#047857; }
        section[data-testid="stSidebar"] { background:#F8FAFC; border-right:1px solid #E4E7EB; }
        [data-testid="stMetricValue"] { color:#1E3A5F; font-family:'Poppins',sans-serif; }
        @media (max-width:640px) {
            .block-container { padding:1rem .7rem 4rem !important; }
            h1 { font-size:1.6rem !important; }
            [data-testid="stMetricValue"] { font-size:1.1rem !important; }
        }
        </style>
        """,
        unsafe_allow_html=True)

    if "logo_b64" not in st.session_state:
        with open(LOGO_PATH, "rb") as f:
            st.session_state["logo_b64"] = base64.b64encode(f.read()).decode()
    if "sig_b64" not in st.session_state:
        st.session_state["sig_b64"] = signature_b64()

    st.markdown(
        "<h1 style='margin-bottom:0'>Apple Reigns Enterprise</h1>"
        "<p style='color:#64748B;margin-top:4px'>Billing suite — invoices, payment receipts, and revenue analytics.</p>",
        unsafe_allow_html=True)

    mode = st.sidebar.radio(
        "Generate", ["🧾 Invoice", "🧾 Payment receipt", "📈 Revenue analytics"],
        index=0)
    st.sidebar.divider()

    if mode.endswith("Invoice"):
        invoice_view()
    elif mode.endswith("receipt"):
        receipt_view()
    else:
        analytics_view()


if __name__ == "__main__":
    main()
