import io
import re
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st

st.set_page_config(page_title="QF | QuoteFast", page_icon="⚡", layout="wide")
st.title("⚡ QF | QuoteFast")
st.caption("From RFQ to Quotation in Minutes — v0.3")
st.info("Demo: use fictional data only. Suggested product matches must be reviewed before quoting.")


def normalize(text):
    return re.sub(r"\s+", " ", str(text).lower().strip())


def find_col(columns, options):
    names = {normalize(c): c for c in columns}
    for option in options:
        if option in names:
            return names[option]
    for col in columns:
        if any(option in normalize(col) for option in options):
            return col
    return None


def read_table(file):
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(file.getvalue()), dtype=str).fillna("")
    if name.endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(file.getvalue()), engine="openpyxl", dtype=str).fillna("")
    raise ValueError("Upload an Excel (.xlsx) or CSV file.")


def parse_catalogue(file):
    df = read_table(file)
    sku_col = find_col(df.columns, ["sku", "product code", "item code", "part number", "code"])
    desc_col = find_col(df.columns, ["description", "product name", "item name", "product"])
    price_col = find_col(df.columns, ["unit price", "price", "rate"])
    if not all([sku_col, desc_col, price_col]):
        raise ValueError("Catalogue needs SKU, Description and Price columns.")
    out = pd.DataFrame({
        "SKU": df[sku_col].astype(str).str.strip().str.upper(),
        "Description": df[desc_col].astype(str).str.strip(),
        "Unit Price": pd.to_numeric(df[price_col].astype(str).str.replace(",", "", regex=False), errors="coerce"),
    })
    return out[(out["SKU"] != "") & (out["Description"] != "")].reset_index(drop=True)


def parse_rfq(file):
    name = file.name.lower()
    if name.endswith((".xlsx", ".csv")):
        # Header-aware parsing first, then headerless fallback for sample RFQs.
        df = read_table(file)
        sku_col = find_col(df.columns, ["sku", "product code", "item code", "part number", "code"])
        desc_col = find_col(df.columns, ["description", "item name", "product", "item description"])
        qty_col = find_col(df.columns, ["quantity", "qty", "requested quantity"])
        if qty_col and (sku_col or desc_col):
            rows = []
            for _, row in df.iterrows():
                sku = str(row[sku_col]).strip().upper() if sku_col else ""
                desc = str(row[desc_col]).strip() if desc_col else ""
                qty = pd.to_numeric(str(row[qty_col]).replace(",", ""), errors="coerce")
                if pd.notna(qty) and qty > 0 and (sku or desc):
                    rows.append({"Requested SKU": sku, "Requested Description": desc, "Quantity": float(qty)})
            if rows:
                return pd.DataFrame(rows)
        if name.endswith(".xlsx"):
            raw = pd.read_excel(io.BytesIO(file.getvalue()), header=None, engine="openpyxl", dtype=str).fillna("")
        else:
            raw = pd.read_csv(io.BytesIO(file.getvalue()), header=None, dtype=str).fillna("")
        return parse_lines("\n".join(",".join(map(str, row)) for row in raw.values.tolist()))
    if name.endswith(".pdf"):
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(file.getvalue()))
        return parse_lines("\n".join(page.extract_text() or "" for page in reader.pages))
    if name.endswith(".docx"):
        from docx import Document
        doc = Document(io.BytesIO(file.getvalue()))
        lines = [p.text for p in doc.paragraphs]
        lines += [",".join(c.text for c in row.cells) for t in doc.tables for row in t.rows]
        return parse_lines("\n".join(lines))
    raise ValueError("Unsupported RFQ format.")


def parse_lines(text):
    rows = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in re.split(r"[,;\t]", line) if p.strip()]
        if len(parts) < 2:
            parts = line.rsplit(" ", 1)
        if len(parts) < 2:
            continue
        qty = pd.to_numeric(parts[-1], errors="coerce")
        if pd.isna(qty) or qty <= 0:
            continue
        item = " ".join(parts[:-1]).strip()
        if not item or normalize(item) in ("sku", "description", "item"):
            continue
        sku = item.upper() if re.fullmatch(r"SKU[\w-]+", item, flags=re.I) else ""
        rows.append({"Requested SKU": sku, "Requested Description": "" if sku else item, "Quantity": float(qty)})
    return pd.DataFrame(rows, columns=["Requested SKU", "Requested Description", "Quantity"])


def suggest_match(sku, description, catalogue):
    exact = catalogue[catalogue["SKU"] == sku.upper()] if sku else catalogue.iloc[0:0]
    if len(exact) == 1:
        return exact.iloc[0], "Exact SKU", 1.0
    if len(exact) > 1:
        return None, "Duplicate SKU — review", 0.0
    if not description:
        return None, "No match", 0.0
    scores = catalogue["Description"].map(lambda x: SequenceMatcher(None, normalize(description), normalize(x)).ratio())
    if scores.empty:
        return None, "No match", 0.0
    idx = scores.idxmax()
    score = float(scores.loc[idx])
    if score >= 0.65:
        return catalogue.loc[idx], "Suggested — verify", score
    return None, "No reliable match", score


def excel_bytes(df, currency):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Draft Quotation")
        ws = writer.sheets["Draft Quotation"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        widths = {"A": 20, "B": 38, "C": 12, "D": 20, "E": 38, "F": 17, "G": 15, "H": 17, "I": 24}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
        ws.cell(row=len(df) + 3, column=7, value=f"Currency: {currency}")
        ws.cell(row=len(df) + 4, column=7, value="DRAFT — REVIEW REQUIRED")
    return output.getvalue()

catalogue_file = st.file_uploader("1. Upload product catalogue (CSV or Excel)", type=["csv", "xlsx"])
rfq_file = st.file_uploader("2. Upload customer RFQ (Excel, CSV, PDF, Word)", type=["xlsx", "csv", "pdf", "docx"])
rfq_text = st.text_area("Or paste RFQ lines (e.g. SKU001, 10)", height=100)
currency = st.selectbox("Currency", ["USD", "EUR", "GBP"])

if st.button("Generate Draft Quotation", type="primary"):
    if not catalogue_file:
        st.warning("Upload the product catalogue first.")
    elif not rfq_file and not rfq_text.strip():
        st.warning("Upload or paste a customer RFQ.")
    else:
        try:
            catalogue = parse_catalogue(catalogue_file)
            rfq = parse_rfq(rfq_file) if rfq_file else parse_lines(rfq_text)
            if rfq.empty:
                st.error("No RFQ lines found. Use columns SKU/Description and Quantity, or lines like SKU001, 10.")
            else:
                rows = []
                for _, item in rfq.iterrows():
                    product, status, score = suggest_match(item["Requested SKU"], item["Requested Description"], catalogue)
                    rows.append({
                        "Requested SKU": item["Requested SKU"],
                        "Requested Description": item["Requested Description"],
                        "Quantity": item["Quantity"],
                        "Matched SKU": product["SKU"] if product is not None else "",
                        "Matched Description": product["Description"] if product is not None else "",
                        "Unit Price": float(product["Unit Price"]) if product is not None and pd.notna(product["Unit Price"]) else None,
                        "Confidence": round(score * 100),
                        "Status": status,
                    })
                st.session_state["draft"] = pd.DataFrame(rows)
        except Exception as exc:
            st.error(f"Could not process input: {exc}")

if "draft" in st.session_state:
    st.subheader("3. Review and edit draft quotation")
    st.caption("Edit the matched SKU, description, quantity or price. Description matches are suggestions only.")
    edited = st.data_editor(
        st.session_state["draft"],
        hide_index=True,
        num_rows="dynamic",
        disabled=["Requested SKU", "Requested Description", "Confidence", "Status"],
        key="quotation_editor",
        use_container_width=True,
    )
    edited["Quantity"] = pd.to_numeric(edited["Quantity"], errors="coerce")
    edited["Unit Price"] = pd.to_numeric(edited["Unit Price"], errors="coerce")
    edited["Total"] = edited["Quantity"] * edited["Unit Price"]
    unresolved = edited["Matched SKU"].fillna("").astype(str).str.strip().eq("") | edited["Unit Price"].isna()
    st.metric("Draft subtotal (priced lines only)", f"{currency} {edited.loc[~unresolved, 'Total'].sum():,.2f}")
    if unresolved.any():
        st.warning(f"{int(unresolved.sum())} line(s) remain unmatched or unpriced. Total is incomplete.")
    st.warning("Review technical specifications, quantities, prices and proposed matches before sending a quotation.")
    st.download_button("Download Draft Quotation (Excel)", excel_bytes(edited, currency), "QuoteFast_Draft_Quotation.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.caption("QuoteFast v0.3 — prototype. Text similarity is not AI technical verification.")
