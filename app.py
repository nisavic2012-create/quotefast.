
import io
import re
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st

st.set_page_config(page_title="QF | QuoteFast", page_icon="⚡", layout="wide")
st.title("⚡ QF | QuoteFast")
st.caption("From RFQ to Quotation in Minutes — v0.4")
st.info("Prototype: use fictional data. Technical suitability and prices require human approval.")


def norm(value):
    return re.sub(r"\s+", " ", str(value).lower().strip())


def colname(columns, choices):
    mapping = {norm(c): c for c in columns}
    for choice in choices:
        if choice in mapping:
            return mapping[choice]
    for c in columns:
        if any(choice in norm(c) for choice in choices):
            return c
    return None


def read_tabular(file, header=0):
    data = io.BytesIO(file.getvalue())
    if file.name.lower().endswith(".xlsx"):
        return pd.read_excel(data, engine="openpyxl", dtype=str, header=header).fillna("")
    return pd.read_csv(data, dtype=str, header=header, encoding="utf-8-sig").fillna("")


def parse_catalogue(file):
    df = read_tabular(file)
    sku = colname(df.columns, ["sku", "product code", "item code", "part number", "code"])
    desc = colname(df.columns, ["description", "product name", "item name", "product"])
    price = colname(df.columns, ["unit price", "price", "rate"])
    if not all([sku, desc, price]):
        raise ValueError("Catalogue requires SKU, Description, and Price columns.")
    out = pd.DataFrame({
        "SKU": df[sku].astype(str).str.strip().str.upper(),
        "Description": df[desc].astype(str).str.strip(),
        "Unit Price": pd.to_numeric(df[price].astype(str).str.replace(",", "", regex=False), errors="coerce"),
    })
    return out[(out["SKU"] != "") & (out["Description"] != "")].reset_index(drop=True)


QTY_RE = re.compile(r"^(.*?)\s*[,;\t]\s*(\d+(?:\.\d+)?)\s*(?:pcs?|pieces?|ea|each|units?|nos?\.?|sets?|m|meters?)?\s*$", re.I)
QTY_SPACE_RE = re.compile(r"^(.*?)\s+(\d+(?:\.\d+)?)\s*(?:pcs?|pieces?|ea|each|units?|nos?\.?|sets?|m|meters?)\s*$", re.I)
SKU_RE = re.compile(r"^SKU[\w-]+$", re.I)


def parse_lines(text):
    records = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        m = QTY_RE.match(line) or QTY_SPACE_RE.match(line)
        if not m:
            continue
        item = m.group(1).strip(" ,;\t")
        qty = float(m.group(2))
        if not item or qty <= 0 or norm(item) in ("sku", "description", "item"):
            continue
        sku = item.upper() if SKU_RE.fullmatch(item) else ""
        records.append({"Requested SKU": sku, "Requested Description": "" if sku else item, "Quantity": qty})
    return pd.DataFrame(records, columns=["Requested SKU", "Requested Description", "Quantity"])


def parse_rfq(file):
    suffix = file.name.lower()
    if suffix.endswith((".xlsx", ".csv")):
        df = read_tabular(file)
        sku_col = colname(df.columns, ["sku", "product code", "item code", "part number", "code"])
        desc_col = colname(df.columns, ["description", "item description", "item name", "product"])
        qty_col = colname(df.columns, ["quantity", "qty", "requested quantity"])
        if qty_col and (sku_col or desc_col):
            records = []
            for _, row in df.iterrows():
                sku = str(row[sku_col]).strip().upper() if sku_col else ""
                desc = str(row[desc_col]).strip() if desc_col else ""
                qty_text = str(row[qty_col]).strip().replace(",", "")
                qty_match = re.match(r"^(\d+(?:\.\d+)?)", qty_text)
                if qty_match and float(qty_match.group(1)) > 0 and (sku or desc):
                    records.append({"Requested SKU": sku, "Requested Description": desc, "Quantity": float(qty_match.group(1))})
            if records:
                return pd.DataFrame(records)
        raw = read_tabular(file, header=None)
        return parse_lines("\n".join(",".join(str(cell) for cell in row if str(cell).strip()) for row in raw.values.tolist()))
    if suffix.endswith(".pdf"):
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(file.getvalue()))
        return parse_lines("\n".join(page.extract_text() or "" for page in reader.pages))
    if suffix.endswith(".docx"):
        from docx import Document
        doc = Document(io.BytesIO(file.getvalue()))
        lines = [p.text for p in doc.paragraphs]
        lines.extend(",".join(c.text for c in row.cells) for table in doc.tables for row in table.rows)
        return parse_lines("\n".join(lines))
    raise ValueError("Unsupported RFQ file format.")


FAMILIES = ["ball valve", "gate valve", "check valve", "globe valve", "flange", "pipe", "elbow", "tee", "reducer"]


def specs(text):
    t = norm(text).replace('″', '"').replace('”', '"')
    family = next((x for x in FAMILIES if x in t), None)
    if family is None and "valve" in t:
        family = "valve"
    size = re.search(r"\b(\d+(?:\.\d+)?)\s*(?:inch|inches|in\b|\")", t)
    pressure = re.search(r"\b(?:class|cl|#)\s*[-:]?\s*(150|300|600|900|1500|2500)\b", t)
    material = next((x for x in ["stainless steel", "carbon steel", "a105", "a182", "316", "304"] if x in t), None)
    return {"family": family, "size": size.group(1) if size else None,
            "class": pressure.group(1) if pressure else None, "material": material}


def match_product(sku, description, catalogue):
    if sku:
        exact = catalogue[catalogue["SKU"] == sku.upper()]
        if len(exact) == 1:
            return exact.iloc[0], "Exact SKU — verify specs", 100, "SKU found; verify technical specifications"
        if len(exact) > 1:
            return None, "Duplicate SKU — review", 0, "Catalogue has duplicate SKUs"
        return None, "SKU not found", 0, "No exact catalogue SKU"
    if not description or catalogue.empty:
        return None, "Not found", 0, "No description or catalogue"
    requested = specs(description)
    candidates = []
    for _, product in catalogue.iterrows():
        offered = specs(product["Description"])
        if requested["family"] and offered["family"] != requested["family"]:
            continue
        conflict = [k for k in ("size", "class", "material")
                    if requested[k] and offered[k] and requested[k] != offered[k]]
        if conflict:
            continue
        score = SequenceMatcher(None, norm(description), norm(product["Description"])).ratio()
        candidates.append((score, product, offered))
    if not candidates:
        return None, "No suitable candidate", 0, "Product family or specified attributes conflict"
    score, product, offered = max(candidates, key=lambda x: x[0])
    if score < 0.40:
        return None, "No reliable candidate", round(score * 100), "Low text similarity"
    unspecified = [k for k in ("size", "class", "material") if requested[k] and not offered[k]]
    if unspecified:
        return product, "Suggested — specs unverified", round(score * 100), "Catalogue omits: " + ", ".join(unspecified)
    return product, "Suggested — verify", round(score * 100), "Text similarity only; human verification required"


def excel_export(df, currency):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Draft Quotation")
        sheet = writer.sheets["Draft Quotation"]
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in sheet.columns:
            letter = column[0].column_letter
            sheet.column_dimensions[letter].width = min(55, max(15, max(len(str(c.value or "")) for c in column) + 2))
        sheet.cell(len(df) + 3, 1, f"Currency: {currency}")
        sheet.cell(len(df) + 4, 1, "DRAFT — NOT TECHNICALLY APPROVED")
    return output.getvalue()


catalogue_file = st.file_uploader("1. Upload product catalogue (CSV or Excel)", type=["csv", "xlsx"])
rfq_file = st.file_uploader("2. Upload customer RFQ (Excel, CSV, PDF, Word)", type=["xlsx", "csv", "pdf", "docx"])
rfq_text = st.text_area("Or paste RFQ lines", height=125, placeholder="Gate valve, 4 inch, Class 300, 5 pcs")
currency = st.selectbox("Currency", ["USD", "EUR", "GBP"])

input_signature = (catalogue_file.getvalue() if catalogue_file else None,
                   rfq_file.getvalue() if rfq_file else None, rfq_text, currency)
if st.session_state.get("input_signature") != input_signature:
    st.session_state["input_signature"] = input_signature
    st.session_state.pop("draft", None)
    st.session_state.pop("quotation_editor", None)

if st.button("Generate Draft Quotation", type="primary"):
    st.session_state.pop("draft", None)
    st.session_state.pop("quotation_editor", None)
    if not catalogue_file:
        st.warning("Upload a product catalogue first.")
    elif not rfq_file and not rfq_text.strip():
        st.warning("Upload an RFQ or paste RFQ lines.")
    else:
        try:
            catalogue = parse_catalogue(catalogue_file)
            rfq = parse_rfq(rfq_file) if rfq_file else parse_lines(rfq_text)
            if rfq.empty:
                st.error("No RFQ lines recognized. Try: Gate valve, 4 inch, Class 300, 5 pcs")
            else:
                rows = []
                for _, item in rfq.iterrows():
                    product, status, score, note = match_product(item["Requested SKU"], item["Requested Description"], catalogue)
                    rows.append({
                        "Requested SKU": item["Requested SKU"],
                        "Requested Description": item["Requested Description"],
                        "Quantity": item["Quantity"],
                        "Matched SKU": product["SKU"] if product is not None else "",
                        "Matched Description": product["Description"] if product is not None else "",
                        "Unit Price": float(product["Unit Price"]) if product is not None and pd.notna(product["Unit Price"]) else None,
                        "Similarity %": score,
                        "Status": status,
                        "Review Note": note,
                        "Approved": False,
                    })
                st.session_state["draft"] = pd.DataFrame(rows)
        except Exception as exc:
            st.error(f"Unable to process RFQ: {exc}")

if "draft" in st.session_state:
    st.subheader("3. Review and edit draft quotation")
    st.caption("Check specifications and edit proposed matches. Tick Approved only after manual verification.")
    edited = st.data_editor(st.session_state["draft"], hide_index=True, num_rows="fixed",
                            disabled=["Requested SKU", "Requested Description", "Similarity %", "Status", "Review Note"],
                            key="quotation_editor", use_container_width=True)
    edited["Quantity"] = pd.to_numeric(edited["Quantity"], errors="coerce")
    edited["Unit Price"] = pd.to_numeric(edited["Unit Price"], errors="coerce")
    edited["Total"] = edited["Quantity"] * edited["Unit Price"]
    valid = (edited["Approved"].fillna(False).astype(bool)
             & edited["Matched SKU"].fillna("").astype(str).str.strip().ne("")
             & edited["Quantity"].gt(0) & edited["Unit Price"].ge(0))
    st.metric("Approved subtotal", f"{currency} {edited.loc[valid, 'Total'].sum():,.2f}")
    st.caption(f"Suggested priced lines (not approved): {currency} {edited.loc[~valid, 'Total'].sum():,.2f}")
    if (~valid).any():
        st.warning(f"{int((~valid).sum())} line(s) are not approved. Do not issue this quotation yet.")
    st.download_button("Download Draft Quotation (Excel)", excel_export(edited, currency),
                       file_name="QuoteFast_v04_Draft.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.caption("QuoteFast v0.4 — prototype. Similarity is not proof of technical equivalence.")
