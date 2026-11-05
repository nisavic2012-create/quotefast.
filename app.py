
import io
import re
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st

st.set_page_config(page_title='QF | QuoteFast', page_icon='⚡', layout='wide')
st.title('⚡ QF | QuoteFast')
st.caption('From RFQ to Quotation in Minutes — v0.5')
st.warning('Prototype: all technical matches, quantities and prices require human verification.')


def norm(v):
    return re.sub(r'\s+', ' ', str(v).lower().strip())


def colname(columns, choices):
    lookup = {norm(c): c for c in columns}
    for choice in choices:
        if choice in lookup:
            return lookup[choice]
    for c in columns:
        if any(choice in norm(c) for choice in choices):
            return c
    return None


def read_tabular(file, header=0):
    data = io.BytesIO(file.getvalue())
    if file.name.lower().endswith('.xlsx'):
        return pd.read_excel(data, engine='openpyxl', dtype=str, header=header).fillna('')
    return pd.read_csv(data, dtype=str, header=header, encoding='utf-8-sig').fillna('')


def parse_catalogue(file):
    df = read_tabular(file)
    sku = colname(df.columns, ['sku', 'product code', 'item code', 'part number', 'code'])
    desc = colname(df.columns, ['description', 'product name', 'item name', 'product'])
    price = colname(df.columns, ['unit price', 'price', 'rate'])
    if not all([sku, desc, price]):
        raise ValueError('Catalogue requires SKU, Description and Price columns.')
    out = pd.DataFrame({
        'SKU': df[sku].astype(str).str.strip().str.upper(),
        'Description': df[desc].astype(str).str.strip(),
        'Unit Price': pd.to_numeric(df[price].astype(str).str.replace(',', '', regex=False), errors='coerce'),
    })
    return out[(out['SKU'] != '') & (out['Description'] != '')].reset_index(drop=True)


QTY_RE = re.compile(r'^(.*?)\s*[,;\t]\s*(\d+(?:\.\d+)?)\s*(?:pcs?|pieces?|ea|each|units?|nos?\.?|sets?|m|meters?)?\s*$', re.I)
QTY_SPACE_RE = re.compile(r'^(.*?)\s+(\d+(?:\.\d+)?)\s*(?:pcs?|pieces?|ea|each|units?|nos?\.?|sets?|m|meters?)\s*$', re.I)
SKU_RE = re.compile(r'^SKU[\w-]+$', re.I)


def parse_lines(text):
    records = []
    for raw in text.splitlines():
        m = QTY_RE.match(raw.strip()) or QTY_SPACE_RE.match(raw.strip())
        if not m:
            continue
        item = m.group(1).strip(' ,;\t')
        qty = float(m.group(2))
        if not item or qty <= 0 or norm(item) in ('sku', 'description', 'item'):
            continue
        sku = item.upper() if SKU_RE.fullmatch(item) else ''
        records.append({'Requested SKU': sku, 'Requested Description': '' if sku else item, 'Quantity': qty})
    return pd.DataFrame(records, columns=['Requested SKU', 'Requested Description', 'Quantity'])


def parse_rfq(file):
    suffix = file.name.lower()
    if suffix.endswith(('.xlsx', '.csv')):
        df = read_tabular(file)
        sku = colname(df.columns, ['sku', 'product code', 'item code', 'part number', 'code'])
        desc = colname(df.columns, ['description', 'item description', 'item name', 'product'])
        qty = colname(df.columns, ['quantity', 'qty', 'requested quantity'])
        if qty and (sku or desc):
            records = []
            for _, row in df.iterrows():
                s = str(row[sku]).strip().upper() if sku else ''
                d = str(row[desc]).strip() if desc else ''
                m = re.match(r'^\s*(\d+(?:\.\d+)?)', str(row[qty]).replace(',', ''))
                if m and float(m.group(1)) > 0 and (s or d):
                    records.append({'Requested SKU': s, 'Requested Description': d, 'Quantity': float(m.group(1))})
            if records:
                return pd.DataFrame(records)
        raw = read_tabular(file, header=None)
        return parse_lines('\n'.join(','.join(str(x) for x in row if str(x).strip()) for row in raw.values.tolist()))
    if suffix.endswith('.pdf'):
        from pypdf import PdfReader
        return parse_lines('\n'.join(page.extract_text() or '' for page in PdfReader(io.BytesIO(file.getvalue())).pages))
    if suffix.endswith('.docx'):
        from docx import Document
        doc = Document(io.BytesIO(file.getvalue()))
        lines = [p.text for p in doc.paragraphs]
        lines += [','.join(c.text for c in row.cells) for table in doc.tables for row in table.rows]
        return parse_lines('\n'.join(lines))
    raise ValueError('Unsupported RFQ format')


FAMILIES = ['ball valve', 'gate valve', 'check valve', 'globe valve', 'flange', 'pipe', 'elbow', 'tee', 'reducer']


def specs(text):
    t = norm(text).replace('″', '"').replace('”', '"').replace('×', 'x')
    family = next((f for f in FAMILIES if re.search(r'\b' + re.escape(f) + r'\b', t)), None)
    if family is None and re.search(r'\bvalve\b', t):
        family = 'valve'
    size = re.search(r'\b(\d+(?:\.\d+)?)\s*(?:inch(?:es)?\b|in\b|")', t)
    dn = re.search(r'\bdn\s*(\d{2,4})\b', t)
    odwt = re.search(r'\b(\d{2,4}(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*(?:mm)?\b', t)
    wt = re.search(r'\b(?:wt|wall thickness|wall|thk)\s*[:=]?\s*(\d+(?:\.\d+)?)\s*mm\b', t)
    pressure = re.search(r'\b(?:class|cl)\s*[-:]?\s*(150|300|600|900|1500|2500)\b|(?<!\w)(150|300|600|900|1500|2500)\s*#', t)
    materials = ['a105n', 'a105', 'a182', '316l', '316ss', '316', '304', 'x42', 'st20', 'carbon steel', 'stainless steel']
    material = next((m for m in materials if re.search(r'(?<!\w)' + re.escape(m) + r'(?!\w)', t)), None)
    standards = re.findall(r'\b(?:api\s*5l|api\s*6a|asme\s*b16\.5|asme\s*b16\.9|astm\s*a\d+[a-z]?)\b', t)
    coating = next((c for c in ['3lpe', '3lp', 'fbe'] if re.search(r'\b' + c + r'\b', t)), None)
    return {'family': family, 'size_in': size.group(1) if size else None,
            'dn': dn.group(1) if dn else None,
            'od_mm': odwt.group(1) if odwt else None,
            'wt_mm': (odwt.group(2) if odwt else wt.group(1) if wt else None),
            'class': next((g for g in pressure.groups() if g), None) if pressure else None,
            'material': material, 'standard': ', '.join(sorted(set(standards))) or None,
            'coating': coating}


def equivalent(a, b):
    if a is None or b is None:
        return True
    try:
        return abs(float(a) - float(b)) < 0.001
    except (ValueError, TypeError):
        return a == b


def assess(request, offered):
    conflicts, missing = [], []
    for key, label in [('family', 'product type'), ('size_in', 'inch size'), ('dn', 'DN'),
                       ('od_mm', 'outside diameter'), ('wt_mm', 'wall thickness'),
                       ('class', 'pressure class'), ('material', 'material'),
                       ('standard', 'standard'), ('coating', 'coating')]:
        if request[key] is None:
            continue
        if offered[key] is None:
            missing.append(label)
        elif not equivalent(request[key], offered[key]):
            conflicts.append(f'{label}: requested {request[key]}, offered {offered[key]}')
    return conflicts, missing


def match_product(sku, description, catalogue):
    if sku:
        matches = catalogue[catalogue['SKU'] == sku.upper()]
        if len(matches) != 1:
            return None, 'Duplicate SKU' if len(matches) > 1 else 'SKU not found', 0, 'Manual review required'
        product = matches.iloc[0]
        conflicts, missing = assess(specs(description), specs(product['Description'])) if description else ([], [])
        if conflicts:
            return None, 'Technical conflict', 0, '; '.join(conflicts)
        return product, 'Exact SKU — review', 100, ('Missing: ' + ', '.join(missing)) if missing else 'Verify specifications manually'
    if not description:
        return None, 'No description', 0, 'Manual review required'
    requested = specs(description)
    candidates = []
    conflicts_seen = []
    for _, product in catalogue.iterrows():
        offered = specs(product['Description'])
        conflicts, missing = assess(requested, offered)
        if conflicts:
            conflicts_seen.extend(conflicts)
            continue
        score = SequenceMatcher(None, norm(description), norm(product['Description'])).ratio()
        candidates.append((score, len(missing), product, missing))
    if not candidates:
        return None, 'No compatible candidate', 0, '; '.join(sorted(set(conflicts_seen))[:3]) or 'No candidates'
    score, _, product, missing = max(candidates, key=lambda x: (x[0], -x[1]))
    if score < 0.40:
        return None, 'Low similarity', round(score * 100), 'No reliable text match'
    if missing:
        return product, 'Suggested — specs missing', round(score * 100), 'Catalogue missing: ' + ', '.join(missing)
    return product, 'Suggested — review', round(score * 100), 'Technical compatibility not certified'


def excel_export(df, currency):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Draft Quotation')
        sheet = writer.sheets['Draft Quotation']
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions
        for column in sheet.columns:
            letter = column[0].column_letter
            sheet.column_dimensions[letter].width = min(55, max(15, max(len(str(c.value or '')) for c in column) + 2))
        sheet.cell(len(df) + 3, 1, f'Currency: {currency}')
        sheet.cell(len(df) + 4, 1, 'DRAFT — TECHNICAL VERIFICATION REQUIRED')
    return output.getvalue()


catalogue_file = st.file_uploader('1. Upload catalogue (CSV or Excel)', type=['csv', 'xlsx'])
rfq_file = st.file_uploader('2. Upload RFQ (Excel, CSV, PDF, Word)', type=['xlsx', 'csv', 'pdf', 'docx'])
rfq_text = st.text_area('Or paste RFQ lines', height=140, placeholder='Gate valve, 4 inch, Class 300, 5 pcs')
currency = st.selectbox('Currency', ['USD', 'EUR', 'GBP'])

signature = (catalogue_file.getvalue() if catalogue_file else None,
             rfq_file.getvalue() if rfq_file else None, rfq_text, currency)
if st.session_state.get('signature') != signature:
    st.session_state['signature'] = signature
    st.session_state.pop('draft', None)
    st.session_state.pop('editor', None)

if st.button('Generate Draft Quotation', type='primary'):
    st.session_state.pop('draft', None)
    st.session_state.pop('editor', None)
    if not catalogue_file:
        st.warning('Upload a catalogue first.')
    elif not rfq_file and not rfq_text.strip():
        st.warning('Upload an RFQ or paste RFQ lines.')
    else:
        try:
            catalogue = parse_catalogue(catalogue_file)
            rfq = parse_rfq(rfq_file) if rfq_file else parse_lines(rfq_text)
            if rfq.empty:
                st.error('No RFQ lines recognized. Example: Gate valve, 4 inch, Class 300, 5 pcs')
            else:
                rows = []
                for _, item in rfq.iterrows():
                    product, status, score, note = match_product(item['Requested SKU'], item['Requested Description'], catalogue)
                    rows.append({'Requested SKU': item['Requested SKU'],
                                 'Requested Description': item['Requested Description'],
                                 'Quantity': item['Quantity'],
                                 'Matched SKU': product['SKU'] if product is not None else '',
                                 'Matched Description': product['Description'] if product is not None else '',
                                 'Unit Price': float(product['Unit Price']) if product is not None and pd.notna(product['Unit Price']) else None,
                                 'Similarity %': score, 'Status': status,
                                 'Review Note': note, 'Approved': False})
                st.session_state['draft'] = pd.DataFrame(rows)
        except Exception as exc:
            st.error(f'Unable to process RFQ: {exc}')

if 'draft' in st.session_state:
    st.subheader('3. Review and edit draft quotation')
    st.caption('Technical matches are suggestions. Approve each line only after checking original specifications.')
    edited = st.data_editor(st.session_state['draft'], hide_index=True, num_rows='fixed',
                            disabled=['Requested SKU', 'Requested Description', 'Similarity %', 'Status', 'Review Note'],
                            key='editor', use_container_width=True)
    edited['Quantity'] = pd.to_numeric(edited['Quantity'], errors='coerce')
    edited['Unit Price'] = pd.to_numeric(edited['Unit Price'], errors='coerce')
    edited['Total'] = edited['Quantity'] * edited['Unit Price']
    valid = (edited['Approved'].fillna(False).astype(bool)
             & edited['Matched SKU'].fillna('').astype(str).str.strip().ne('')
             & edited['Quantity'].gt(0) & edited['Unit Price'].ge(0))
    st.metric('Approved subtotal', f"{currency} {edited.loc[valid, 'Total'].sum():,.2f}")
    st.caption(f"Unapproved priced subtotal: {currency} {edited.loc[~valid, 'Total'].sum():,.2f}")
    if (~valid).any():
        st.warning(f"{int((~valid).sum())} line(s) remain unapproved.")
    st.download_button('Download Draft Quotation (Excel)', excel_export(edited, currency),
                       file_name='QuoteFast_v05_Draft.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

st.divider()
st.caption('QuoteFast v0.5 — rule-based prototype, not certified technical verification.')

