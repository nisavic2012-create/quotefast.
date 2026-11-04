
import io
import re

import pandas as pd
import streamlit as st

st.set_page_config(page_title="QF | QuoteFast", page_icon="⚡")

st.title("⚡ QF | QuoteFast")
st.caption("Version 0.2 — RFQ File Upload")
st.write("From RFQ to Quotation in Minutes")

catalogue_file = st.file_uploader(
    "1. Upload product catalogue (CSV)",
    type=["csv"],
    key="catalogue"
)

rfq_file = st.file_uploader(
    "2. Upload customer RFQ (Excel, CSV, PDF, Word)",
    type=["xlsx", "xls", "csv", "pdf", "docx"],
    key="rfq"
)

rfq_text = st.text_area(
    "Or paste RFQ text",
    placeholder="SKU001, 10\nSKU002, 5"
)

currency = st.selectbox("Currency", ["USD", "EUR", "GBP"])


def read_rfq(file):
    name = file.name.lower()

    if name.endswith((".xlsx", ".xls")):
        if name.endswith(".xlsx"):
            df = pd.read_excel(file, engine="openpyxl")
        else:
            df = pd.read_excel(file)
        return df.to_csv(index=False, header=False)

    if name.endswith(".csv"):
        return file.getvalue().decode("utf-8-sig")

    if name.endswith(".pdf"):
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(file.getvalue()))
        return "\n".join(
            page.extract_text() or "" for page in reader.pages
        )

    if name.endswith(".docx"):
        from docx import Document
        doc = Document(io.BytesIO(file.getvalue()))
        lines = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                lines.append(
                    ",".join(cell.text for cell in row.cells)
                )
        return "\n".join(lines)

    return ""


def extract_items(text):
    items = []
    for line in text.splitlines():
        match = re.search(
            r"\b(SKU[\w-]+)\s*[,;\t ]+\s*(\d+)\b",
            line,
            re.IGNORECASE
        )
        if match:
            items.append((
                match.group(1).upper(),
                int(match.group(2))
            ))
    return items


if st.button("Generate Quotation", type="primary"):

    if catalogue_file is None:
        st.warning("Please upload the product catalogue.")

    elif rfq_file is None and not rfq_text.strip():
        st.warning("Please upload or paste a customer RFQ.")

    else:
        try:
            catalogue = pd.read_csv(catalogue_file)
            catalogue.columns = [
                str(c).strip().lower()
                for c in catalogue.columns
            ]

            required = {"sku", "description", "price"}
            if not required.issubset(catalogue.columns):
                st.error(
                    "Catalogue needs sku, description, price columns."
                )
                st.stop()

            catalogue["sku"] = (
                catalogue["sku"].astype(str).str.strip().str.upper()
            )
            catalogue["price"] = pd.to_numeric(
                catalogue["price"], errors="coerce"
            )

            text = read_rfq(rfq_file) if rfq_file else rfq_text
            items = extract_items(text)

            if not items:
                st.error(
                    "No SKU and quantity pairs found. "
                    "Try lines like SKU001, 10."
                )
                st.stop()

            rows = []
            for sku, qty in items:
                matches = catalogue[catalogue["sku"] == sku]

                if len(matches) == 1:
                    product = matches.iloc[0]
                    price = product["price"]
                    status = (
                        "Matched" if pd.notna(price)
                        else "Price missing"
                    )
                    description = product["description"]
                elif len(matches) > 1:
                    price = None
                    status = "Duplicate SKU"
                    description = ""
                else:
                    price = None
                    status = "Not found"
                    description = ""

                rows.append({
                    "SKU": sku,
                    "Description": description,
                    "Quantity": qty,
                    "Unit Price": price,
                    "Total": (
                        round(qty * price, 2)
                        if pd.notna(price) else None
                    ),
                    "Status": status
                })

            result = pd.DataFrame(rows)

            st.subheader("Draft Quotation")
            st.dataframe(result, use_container_width=True)

            matched = result[result["Status"] == "Matched"]
            subtotal = matched["Total"].sum()

            st.metric(
                "Matched subtotal",
                f"{currency} {subtotal:,.2f}"
            )

            if len(matched) != len(result):
                st.warning(
                    "Some items need review before quotation."
                )

            st.download_button(
                "Download Quotation CSV",
                result.to_csv(index=False),
                file_name="QuoteFast_Quotation.csv",
                mime="text/csv"
            )

        except ImportError as error:
            st.error(f"Missing Python library: {error}")
            st.info(
                "Check requirements.txt and deployment logs."
            )
        except Exception as error:
            st.error(f"Unable to process RFQ: {error}")

st.divider()
st.caption(
    "QuoteFast v0.2 | Prototype — "
    "Verify all quantities and prices before issuing."
)
