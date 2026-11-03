
import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="QF | QuoteFast", page_icon="⚡")

st.title("QF | QuoteFast")
st.caption("From RFQ to Quotation in Minutes")

st.write("Upload your product catalogue and paste a customer RFQ.")

catalogue_file = st.file_uploader(
    "Upload product catalogue (CSV)",
    type=["csv"]
)

rfq_text = st.text_area(
    "Paste customer RFQ",
    placeholder="SKU001, 10\nSKU002, 5",
    height=180
)

currency = st.selectbox(
    "Currency",
    ["USD", "EUR", "GBP"]
)

if st.button("Generate Quotation", type="primary"):
    if catalogue_file is None or not rfq_text.strip():
        st.warning("Please upload a catalogue and enter an RFQ.")
    else:
        try:
            catalogue = pd.read_csv(catalogue_file)
            catalogue.columns = [
                str(c).strip().lower() for c in catalogue.columns
            ]

            required = {"sku", "description", "price"}
            if not required.issubset(catalogue.columns):
                st.error(
                    "Catalogue must contain: sku, description, price"
                )
            else:
                catalogue["sku"] = (
                    catalogue["sku"].astype(str).str.strip().str.upper()
                )
                catalogue["price"] = pd.to_numeric(
                    catalogue["price"], errors="coerce"
                )

                items = []
                for line in rfq_text.splitlines():
                    match = re.fullmatch(
                        r"\s*([^,]+)\s*,\s*(\d+)\s*", line
                    )
                    if match:
                        sku = match.group(1).strip().upper()
                        qty = int(match.group(2))
                        found = catalogue[catalogue["sku"] == sku]

                        if len(found) == 1:
                            product = found.iloc[0]
                            price = product["price"]
                            status = (
                                "Matched" if pd.notna(price)
                                else "Price missing"
                            )
                            description = product["description"]
                        elif len(found) > 1:
                            price = None
                            status = "Duplicate SKU - review"
                            description = ""
                        else:
                            price = None
                            status = "Not found - review"
                            description = ""

                        items.append({
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

                if not items:
                    st.error(
                        "No valid RFQ lines found. Use SKU, quantity."
                    )
                else:
                    result = pd.DataFrame(items)
                    st.subheader("Draft Quotation")
                    st.dataframe(result, use_container_width=True)

                    matched = result[result["Status"] == "Matched"]
                    subtotal = matched["Total"].sum()

                    st.metric(
                        "Matched items subtotal",
                        f"{currency} {subtotal:,.2f}"
                    )

                    if len(matched) != len(result):
                        st.warning(
                            "Some items require manual review. "
                            "The subtotal excludes unmatched items."
                        )

                    output = io.StringIO()
                    result.to_csv(output, index=False)

                    st.download_button(
                        "Download Draft Quotation (CSV)",
                        output.getvalue(),
                        file_name="quotefast_draft.csv",
                        mime="text/csv"
                    )

        except Exception as exc:
            st.error(f"Unable to process file: {exc}")

st.divider()
st.caption(
    "Prototype v0.1 — all prices and product matches "
    "must be verified before issuing a quotation."
)
