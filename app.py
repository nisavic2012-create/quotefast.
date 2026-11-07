import io
import re
from difflib import SequenceMatcher
from datetime import date, timedelta

import pandas as pd
import streamlit as st

st.set_page_config(page_title='QF | QuoteFast', page_icon='⚡', layout='wide')
st.title('⚡ QF | QuoteFast')
st.caption('From RFQ to Quotation in Minutes — v0.9.2')
st.warning('Prototype: rule-based suggestions are not technical certification. Verify all specifications manually.')


def norm(value):
    return re.sub(r'\s+', ' ', str(value).lower().strip())


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
FIELDS = [('family', 'Product type'), ('size_in', 'Inch size'), ('dn', 'DN'),
          ('od_mm', 'Outside diameter (mm)'), ('wt_mm', 'Wall thickness (mm)'),
          ('class', 'Pressure class'), ('material', 'Material'),
          ('standard', 'Standard'), ('coating', 'Coating')]


def specs(text):
    t = norm(text).replace('″', '"').replace('”', '"').replace('×', 'x')
    family = next((f for f in FAMILIES if re.search(r'\b' + re.escape(f) + r'\b', t)), None)
    if family is None and re.search(r'\bvalve\b', t):
        family = 'valve'
    size = re.search(r'\b(\d+(?:\.\d+)?)\s*(?:inch(?:es)?\b|in\b|")', t)
    dn = re.search(r'\bdn\s*(\d{2,4})\b', t)
    odwt = re.search(r'\b(?:od\s*)?(\d{2,4}(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*(?:mm)?\b', t)
    wt = re.search(r'\b(?:wt|wall thickness|wall|thk)\s*[:=]?\s*(\d+(?:\.\d+)?)\s*mm\b', t)
    pressure = re.search(r'\b(?:class|cl)\s*[-:]?\s*(150|300|600|900|1500|2500)\b|(?<!\w)(150|300|600|900|1500|2500)\s*#', t)
    materials = ['a105n', 'a105', 'a182', '316l', '316ss', '316', '304', 'x42', 'st20', 'carbon steel', 'stainless steel']
    material = next((m for m in materials if re.search(r'(?<!\w)' + re.escape(m) + r'(?!\w)', t)), None)
    standards = re.findall(r'\b(?:api\s*5l|api\s*6a|asme\s*b16\.5|asme\s*b16\.9|astm\s*a\d+[a-z]?)\b', t)
    coating = next((c for c in ['3lpe', '3lp', 'fbe'] if re.search(r'\b' + c + r'\b', t)), None)
    return {'family': family, 'size_in': size.group(1) if size else None,
            'dn': dn.group(1) if dn else None, 'od_mm': odwt.group(1) if odwt else None,
            'wt_mm': odwt.group(2) if odwt else wt.group(1) if wt else None,
            'class': next((g for g in pressure.groups() if g), None) if pressure else None,
            'material': material, 'standard': ', '.join(sorted(set(standards))) or None,
            'coating': coating}


def equivalent(a, b):
    if a is None or b is None:
        return False
    try:
        return abs(float(a) - float(b)) < 0.001
    except (ValueError, TypeError):
        return a == b


def compliance(request_text, offered_text):
    requested, offered = specs(request_text), specs(offered_text)
    details, conflicts, missing = [], [], []
    for key, label in FIELDS:
        r, o = requested[key], offered[key]
        if r is None:
            state = 'Not specified in RFQ'
        elif o is None:
            state = 'Missing Information'
            missing.append(label)
        elif equivalent(r, o):
            state = 'Matched'
        else:
            state = 'Non-compliant'
            conflicts.append(f'{label}: requested {r}, offered {o}')
        details.append({'Parameter': label, 'Requested': r or '—', 'Offered': o or '—', 'Result': state})
    if conflicts:
        return 'Non-compliant', '; '.join(conflicts), pd.DataFrame(details)
    if missing:
        return 'Missing Information', 'Catalogue missing: ' + ', '.join(missing), pd.DataFrame(details)
    return 'Compatible', 'Recognized parameters match; manual verification required', pd.DataFrame(details)


def general_attributes(description):
    """Conservative detection of explicitly stated general-product attributes."""
    t = norm(description).replace('×', 'x')
    patterns = {
        'Power (W)': r'(?<![\w.])(\d+(?:\.\d+)?)\s*w\b',
        'Voltage (V)': r'(?<![\w.])(\d+(?:\.\d+)?)\s*v\b',
        'Current (A)': r'(?<![\w.])(\d+(?:\.\d+)?)\s*a\b',
        'Capacity (mAh)': r'(?<![\w.])(\d+(?:\.\d+)?)\s*mah\b',
        'Capacity (Ah)': r'(?<![\w.])(\d+(?:\.\d+)?)\s*ah\b',
        'Length (m)': r'(?<![\w.])(\d+(?:\.\d+)?)\s*m\b',
        'Length (mm)': r'(?<![\w.])(\d+(?:\.\d+)?)\s*mm\b',
        'Socket count': r'(?<![\w.])(\d+)\s*sockets?\b',
    }
    return {name: match.group(1) for name, pattern in patterns.items()
            if (match := re.search(pattern, t))}


def general_compliance(request_text, offered_text):
    req, off = norm(request_text), norm(offered_text)
    if not req:
        return 'Missing Information', 'RFQ description missing; manual review required', pd.DataFrame()
    requested, offered = general_attributes(request_text), general_attributes(offered_text)
    details, conflicts, missing = [], [], []
    for label, value in requested.items():
        other = offered.get(label)
        if other is None:
            state = 'Missing Information'
            missing.append(label)
        elif equivalent(value, other):
            state = 'Matched'
        else:
            state = 'Non-compliant'
            conflicts.append(f'{label}: requested {value}, offered {other}')
        details.append({'Parameter': label, 'Requested': value, 'Offered': other or '—', 'Result': state})
    details.insert(0, {'Parameter': 'Full product description',
                       'Requested': request_text, 'Offered': offered_text,
                       'Result': 'Matched' if req == off else 'Review required'})
    if conflicts:
        return 'Non-compliant', '; '.join(conflicts), pd.DataFrame(details)
    if missing:
        return 'Missing Information', 'Offered description missing: ' + ', '.join(missing), pd.DataFrame(details)
    if req != off:
        return 'Missing Information', 'Descriptions differ; manual review required', pd.DataFrame(details)
    return 'Compatible', 'Exact normalized description; manual verification still required', pd.DataFrame(details)



def safe_compliance(request_text, offered_text):
    """Mode-independent safety check. Unrecognized categories require exact descriptions."""
    industrial_req = specs(request_text).get('family')
    industrial_off = specs(offered_text).get('family')
    if industrial_req and industrial_off and industrial_req == industrial_off:
        status, note, details = compliance(request_text, offered_text)
        # General numeric attributes are checked too, even in industrial mode.
        req_attrs = general_attributes(request_text)
        off_attrs = general_attributes(offered_text)
        conflicts = []
        for label, value in req_attrs.items():
            if label in off_attrs and not equivalent(value, off_attrs[label]):
                conflicts.append(f"{label}: requested {value}, offered {off_attrs[label]}")
        if conflicts:
            return 'Non-compliant', '; '.join(conflicts), details
        return status, note, details
    # No recognized matching industrial family: conservative general-product check.
    return general_compliance(request_text, offered_text)

def find_general_product(sku, description, catalogue):
    if sku:
        found = catalogue[catalogue['SKU'] == sku.upper()]
        if len(found) == 1:
            return found.iloc[0], 100, 'Exact SKU — verify specification', False
        return None, 0, 'Duplicate SKU' if len(found) > 1 else 'SKU not found', False
    if not description or catalogue.empty:
        return None, 0, 'No description or catalogue', False
    ranked = [(SequenceMatcher(None, norm(description), norm(row['Description'])).ratio(), row)
              for _, row in catalogue.iterrows()]
    score, product = max(ranked, key=lambda x: x[0])
    if score < 0.55:
        return None, round(score * 100), 'No sufficiently similar candidate', False
    exact = norm(description) == norm(product['Description'])
    return product, round(score * 100), 'Exact description' if exact else 'Candidate — review required', not exact


def find_product(sku, description, catalogue):
    if sku:
        found = catalogue[catalogue['SKU'] == sku.upper()]
        if len(found) == 1:
            return found.iloc[0], 100, 'Exact SKU', False
        return None, 0, 'Duplicate SKU' if len(found) > 1 else 'SKU not found', False
    if not description or catalogue.empty:
        return None, 0, 'No description or catalogue', False
    ranked = []
    requested_family = specs(description)['family']
    for _, product in catalogue.iterrows():
        offered_family = specs(product['Description'])['family']
        # Never suggest a known different product family as an alternative.
        if requested_family and offered_family and requested_family != offered_family:
            continue
        if requested_family and not offered_family:
            continue
        status, note, _ = compliance(description, product['Description'])
        score = SequenceMatcher(None, norm(description), norm(product['Description'])).ratio()
        conflict_count = note.count(';') + 1 if status == 'Non-compliant' else 0
        ranked.append((status, score, conflict_count, product))
    compatible = [x for x in ranked if x[0] != 'Non-compliant' and x[1] >= 0.40]
    if compatible:
        status, score, _, product = max(compatible, key=lambda x: x[1])
        return product, round(score * 100), 'Suggested match', False
    alternatives = [x for x in ranked if x[0] == 'Non-compliant' and x[1] >= 0.40]
    if alternatives:
        status, score, _, product = max(alternatives, key=lambda x: (-x[2], x[1]))
        return product, round(score * 100), 'Alternative — technical deviation', True
    return None, 0, 'No reliable candidate', False


def excel_export(df, matrix, deviations, currency):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Draft Quotation')
        matrix.to_excel(writer, index=False, sheet_name='Compliance Matrix')
        deviations.to_excel(writer, index=False, sheet_name='Technical Deviations')
        for sheet in writer.sheets.values():
            sheet.freeze_panes = 'A2'
            sheet.auto_filter.ref = sheet.dimensions
            for column in sheet.columns:
                letter = column[0].column_letter
                sheet.column_dimensions[letter].width = min(65, max(15, max(len(str(c.value or '')) for c in column) + 2))
        sheet = writer.sheets['Draft Quotation']
        sheet.cell(len(df) + 3, 1, f'Currency: {currency}')
        sheet.cell(len(df) + 4, 1, 'DRAFT — HUMAN TECHNICAL APPROVAL REQUIRED')
        sheet.cell(len(df) + 5, 1, 'NON-COMPLIANT ALTERNATIVES ARE NOT APPROVED QUOTATION ITEMS')
    return output.getvalue()




def commercial_lines(result, mode='Industrial piping & valves'):
    """Independent export gate: never trust the editor's Approved/Compliance flags alone."""
    safe = result['Approved'].eq(True) & result['Compliance'].eq('Compatible')
    safe &= result['Matched SKU'].astype(str).str.strip().ne('')
    safe &= result['Quantity'].gt(0) & result['Unit Price'].ge(0)
    safe &= result['Requested Description'].astype(str).str.strip().ne('')
    for index, row in result.iterrows():
        if not safe.loc[index]:
            continue
        status, _, _ = safe_compliance(str(row['Requested Description']), str(row['Matched Description']))
        if status != 'Compatible':
            safe.loc[index] = False
    return result.loc[safe].copy()


def commercial_excel(lines, meta, currency):
    output = io.BytesIO()
    cols = ['Matched SKU', 'Matched Description', 'Quantity', 'Unit Price', 'Total']
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        lines[cols].rename(columns={'Matched SKU': 'SKU', 'Matched Description': 'Description'}).to_excel(
            writer, sheet_name='Quotation', index=False, startrow=12)
        ws = writer.sheets['Quotation']
        from openpyxl.styles import Font, PatternFill, Alignment
        ws['A1'] = 'QUOTATION — DRAFT / SUBJECT TO FINAL REVIEW'
        ws['A1'].font = Font(size=16, bold=True, color='17365D')
        fields = [('Seller', 'seller'), ('Customer', 'customer'), ('Quotation No.', 'number'),
                  ('Date', 'date'), ('Valid Until', 'valid_until'), ('Delivery', 'delivery'),
                  ('Payment', 'payment'), ('Currency', 'currency'), ('Notes', 'notes')]
        info = dict(meta, currency=currency)
        for row_no, (label, key) in enumerate(fields, 3):
            ws.cell(row_no, 1, label).font = Font(bold=True)
            ws.cell(row_no, 2, str(info.get(key, '') or ''))
        header_row = 13
        for cell in ws[header_row]:
            cell.fill = PatternFill('solid', fgColor='17365D')
            cell.font = Font(bold=True, color='FFFFFF')
        for col, width in {'A': 18, 'B': 62, 'C': 16, 'D': 18, 'E': 20}.items():
            ws.column_dimensions[col].width = width
        for r in range(14, 14 + len(lines)):
            ws.cell(r, 4).number_format = '#,##0.00'
            ws.cell(r, 5).number_format = '#,##0.00'
        total_row = 14 + len(lines)
        ws.cell(total_row, 4, 'TOTAL').font = Font(bold=True)
        ws.cell(total_row, 5, float(lines['Total'].sum())).font = Font(bold=True)
        ws.cell(total_row, 5).number_format = '#,##0.00'
        ws.freeze_panes = 'C14'
        ws.print_options.horizontalCentered = True
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
    return output.getvalue()


def commercial_pdf(lines, meta, currency, logo_bytes=None):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib.utils import ImageReader
    from reportlab.lib.units import mm
    from xml.sax.saxutils import escape
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=(210*mm, 297*mm), leftMargin=16*mm,
                            rightMargin=16*mm, topMargin=17*mm, bottomMargin=17*mm)
    styles = getSampleStyleSheet()
    body = ParagraphStyle('QFBody', parent=styles['BodyText'], fontSize=9, leading=12, spaceAfter=4)
    small = ParagraphStyle('QFSmall', parent=body, fontSize=8, leading=10)
    heading = ParagraphStyle('QFHeading', parent=styles['Heading1'], fontSize=19, textColor=colors.HexColor('#17365D'))
    def para(value, style=body):
        return Paragraph(escape(str(value or '')).replace('\n', '<br/>'), style)
    story = []
    if logo_bytes:
        try:
            img = Image(io.BytesIO(logo_bytes))
            iw, ih = ImageReader(io.BytesIO(logo_bytes)).getSize()
            factor = min(42*mm / iw, 18*mm / ih)
            img.drawWidth, img.drawHeight = iw * factor, ih * factor
            story.append(img)
            story.append(Spacer(1, 5*mm))
        except Exception:
            pass
    story += [para('QUOTATION', heading), para('DRAFT — SUBJECT TO FINAL COMMERCIAL AND TECHNICAL REVIEW', small), Spacer(1, 5*mm)]
    parties = Table([[para('FROM / SELLER', small), para('BILL TO / CUSTOMER', small)],
                     [para(meta['seller']), para(meta['customer'])]],
                    colWidths=[89*mm, 89*mm], hAlign='LEFT')
    parties.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#E9F0F8')),
                                 ('VALIGN',(0,0),(-1,-1),'TOP'),
                                 ('TOPPADDING',(0,0),(-1,-1),7),
                                 ('BOTTOMPADDING',(0,0),(-1,-1),7)]))
    story.extend([parties, Spacer(1, 6*mm)])
    for label, key in [('Quotation No.', 'number'), ('Date', 'date'), ('Valid Until', 'valid_until')]:
        if meta.get(key):
            story.append(para(f'{label}: {meta[key]}', small))
    story.append(Spacer(1, 6*mm))
    widths = [11*mm, 20*mm, 76*mm, 16*mm, 26*mm, 29*mm]
    data = [[para(x, small) for x in ['No.', 'SKU', 'Description', 'Qty', f'Unit ({currency})', f'Total ({currency})']]]
    for no, (_, item) in enumerate(lines.iterrows(), 1):
        data.append([para(no, small), para(item['Matched SKU'], small), para(item['Matched Description'], small),
                     para(f'{float(item["Quantity"]):g}', small),
                     para(f'{float(item["Unit Price"]):,.2f}', small), para(f'{float(item["Total"]):,.2f}', small)])
    data.append(['', '', '', '', para('TOTAL', small), para(f'{float(lines["Total"].sum()):,.2f}', small)])
    table = Table(data, colWidths=widths, repeatRows=1, hAlign='LEFT')
    table.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#17365D')),
                               ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                               ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, colors.HexColor('#F2F5F9')]),
                               ('LINEBELOW', (0,0), (-1,-1), .25, colors.HexColor('#D5DCE5')),
                               ('VALIGN', (0,0), (-1,-1), 'TOP'),
                               ('LEFTPADDING', (0,0), (-1,-1), 5), ('RIGHTPADDING', (0,0), (-1,-1), 5),
                               ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6)]))
    # Header paragraphs need explicit white text; a cell-level TEXTCOLOR doesn't override Paragraph styling.
    for cell in data[0]:
        cell.style = ParagraphStyle('QFHeader', parent=small, textColor=colors.white, fontName='Helvetica-Bold')
    story.extend([table, Spacer(1, 7*mm)])
    for label, key in [('Delivery terms', 'delivery'), ('Payment terms', 'payment'), ('Additional notes', 'notes')]:
        value = str(meta.get(key) or '').strip()
        if value:
            story.append(para(f'{label}: {value}', small))
    story.append(Spacer(1, 4*mm))
    story.append(para('Only manually approved, rule-matched items are included. Verify all specifications before issue.', small))
    doc.build(story)
    return output.getvalue()

industry_mode = st.selectbox('Product matching mode', ['Industrial piping & valves', 'General products (conservative matching)'], help='General mode allows exact-description matching; similar descriptions require manual review.')
st.caption('Matching mode affects suggestions only. Final compliance checks are independent of mode.')
catalogue_file = st.file_uploader('1. Upload catalogue (CSV or Excel)', type=['csv', 'xlsx'])
rfq_file = st.file_uploader('2. Upload RFQ (Excel, CSV, PDF, Word)', type=['xlsx', 'csv', 'pdf', 'docx'])
rfq_text = st.text_area('Or paste RFQ lines', height=140, placeholder='Gate valve, 4 inch, Class 600, 5 pcs')
currency = st.selectbox('Currency', ['USD', 'EUR', 'GBP'])

signature = (catalogue_file.getvalue() if catalogue_file else None,
             rfq_file.getvalue() if rfq_file else None, rfq_text, currency, industry_mode)
if st.session_state.get('signature') != signature:
    st.session_state['signature'] = signature
    for key in ['draft', 'editor', 'catalogue_data']:
        st.session_state.pop(key, None)

if st.button('Generate Draft Quotation', type='primary'):
    for key in ['draft', 'editor', 'catalogue_data']:
        st.session_state.pop(key, None)
    if not catalogue_file:
        st.warning('Upload a catalogue first.')
    elif not rfq_file and not rfq_text.strip():
        st.warning('Upload an RFQ or paste RFQ lines.')
    else:
        try:
            catalogue = parse_catalogue(catalogue_file)
            rfq = parse_rfq(rfq_file) if rfq_file else parse_lines(rfq_text)
            if rfq.empty:
                st.error('No RFQ lines recognized. Example: Gate valve, 4 inch, Class 600, 5 pcs')
            else:
                rows = []
                for _, item in rfq.iterrows():
                    product, score, note, alternative = (find_general_product if industry_mode.startswith('General') else find_product)(item['Requested SKU'], item['Requested Description'], catalogue)
                    rows.append({'Requested SKU': item['Requested SKU'],
                                 'Requested Description': item['Requested Description'],
                                 'Quantity': item['Quantity'],
                                 'Matched SKU': product['SKU'] if product is not None else '',
                                 'Matched Description': product['Description'] if product is not None else '',
                                 'Unit Price': float(product['Unit Price']) if product is not None and pd.notna(product['Unit Price']) else None,
                                 'Similarity %': score, 'Match Note': note, 'Approved': False})
                st.session_state['catalogue_data'] = catalogue
                st.session_state['draft'] = pd.DataFrame(rows)
        except Exception as exc:
            st.error(f'Unable to process RFQ: {exc}')

if 'draft' in st.session_state:
    st.subheader('3. Review and edit draft quotation')
    st.caption('Change Matched SKU to select another product. Only Compatible lines can be approved. General-mode similar descriptions require review.')
    edited = st.data_editor(st.session_state['draft'], hide_index=True, num_rows='fixed',
                            disabled=['Requested SKU', 'Requested Description', 'Matched Description',
                                      'Unit Price', 'Similarity %', 'Match Note'],
                            key='editor', use_container_width=True)
    catalogue = st.session_state['catalogue_data']
    rows, matrix_rows, deviation_rows = [], [], []
    for i, row in edited.iterrows():
        line = row.to_dict()
        sku = str(line.get('Matched SKU') or '').strip().upper()
        matching = catalogue[catalogue['SKU'] == sku] if sku else catalogue.iloc[0:0]
        if len(matching) == 1:
            product = matching.iloc[0]
            line['Matched SKU'] = sku
            line['Matched Description'] = product['Description']
            line['Unit Price'] = product['Unit Price']
            if line['Requested Description']:
                status, note, details = safe_compliance(line['Requested Description'], product['Description'])
            else:
                status, note, details = 'Missing Information', 'RFQ contains SKU only; verify original specification', pd.DataFrame()
        else:
            line['Matched Description'] = ''
            line['Unit Price'] = None
            status = 'No match' if not sku else 'Invalid or duplicate SKU'
            note = 'Choose one unique SKU from the catalogue'
            details = pd.DataFrame()
        line['Compliance'] = status
        line['Compliance Note'] = note
        line['Quantity'] = pd.to_numeric(line['Quantity'], errors='coerce')
        line['Unit Price'] = pd.to_numeric(line['Unit Price'], errors='coerce')
        line['Total'] = line['Quantity'] * line['Unit Price']
        # Even if a user checks Approved, known technical conflicts are forcibly excluded.
        line['Approved'] = bool(line.get('Approved', False)) and status == 'Compatible' and len(matching) == 1
        rows.append(line)
        for _, detail in details.iterrows():
            matrix_rows.append({'RFQ Line': i + 1, 'Matched SKU': sku, **detail.to_dict()})
            if detail['Result'] == 'Non-compliant':
                deviation_rows.append({'RFQ Line': i + 1, 'Alternative SKU': sku,
                                       'Requested Description': line['Requested Description'],
                                       'Alternative Description': line['Matched Description'],
                                       'Parameter': detail['Parameter'],
                                       'Requested': detail['Requested'], 'Offered': detail['Offered'],
                                       'Decision': 'NOT APPROVED — separate engineering review required'})
    result = pd.DataFrame(rows)
    matrix = pd.DataFrame(matrix_rows, columns=['RFQ Line', 'Matched SKU', 'Parameter', 'Requested', 'Offered', 'Result'])
    deviations = pd.DataFrame(deviation_rows, columns=['RFQ Line', 'Alternative SKU', 'Requested Description',
                      'Alternative Description', 'Parameter', 'Requested', 'Offered', 'Decision'])
    st.subheader('4. Technical Compliance Matrix')
    st.dataframe(result[['Requested Description', 'Matched SKU', 'Matched Description', 'Compliance', 'Compliance Note']],
                 hide_index=True, use_container_width=True)
    with st.expander('View parameter-by-parameter comparison', expanded=True):
        st.dataframe(matrix, hide_index=True, use_container_width=True)
    st.subheader('5. Alternative Products & Technical Deviations')
    if deviations.empty:
        st.info('No recognized technical conflicts. Unresolved differences may still require review.')
    else:
        st.error(f'{len(deviations)} technical deviation(s) detected. These alternatives are not approved.')
        st.dataframe(deviations, hide_index=True, use_container_width=True)
    eligible_indices = commercial_lines(result, industry_mode).index
    valid = result.index.isin(eligible_indices)
    st.metric('Approved subtotal', f"{currency} {result.loc[valid, 'Total'].sum():,.2f}")
    st.caption(f"Unapproved priced subtotal (including alternatives): {currency} {result.loc[~valid, 'Total'].sum():,.2f}")
    if (~valid).any():
        st.warning(f"{int((~valid).sum())} line(s) remain unapproved. Do not issue as a final quotation.")
    st.download_button('Download Draft + Compliance + Deviations (Excel)',
                       excel_export(result, matrix, deviations, currency),
                       file_name='QuoteFast_v093_Draft.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


    st.subheader('6. Professional Quotation — v0.9.2')
    st.caption('Customer-facing documents include only approved lines marked Compatible. Technical deviations stay in the internal workbook.')
    with st.expander('Quotation details', expanded=True):
        left, right = st.columns(2)
        with left:
            seller = st.text_area('Seller / company details', placeholder='Company name\nAddress\nEmail', key='qf_seller')
            number = st.text_input('Quotation number', value='QF-001', key='qf_number')
            delivery = st.text_input('Delivery terms', placeholder='e.g. 4–6 weeks, EXW', key='qf_delivery')
            valid_until = st.date_input('Valid until', value=date.today() + timedelta(days=30), key='qf_valid')
        with right:
            customer = st.text_area('Customer details', placeholder='Customer company\nAddress\nContact', key='qf_customer')
            date = st.date_input('Quotation date', key='qf_date')
            payment = st.text_input('Payment terms', placeholder='e.g. 30 days net', key='qf_payment')
            notes = st.text_area('Additional notes', key='qf_notes')
        logo = st.file_uploader('Optional company logo (PNG/JPG)', type=['png', 'jpg', 'jpeg'], key='qf_logo')
    approved_lines = commercial_lines(result, industry_mode)
    st.caption(f'{len(approved_lines)} eligible approved line(s); {len(result)-len(approved_lines)} excluded from customer-facing quotation.')
    if approved_lines.empty:
        st.info('Approve at least one Compatible line above to generate a quotation.')
    else:
        meta = {'seller': seller, 'customer': customer, 'number': number,
                'date': str(date), 'valid_until': str(valid_until),
                'delivery': delivery, 'payment': payment, 'notes': notes}
        st.metric('Customer quotation total', f"{currency} {approved_lines['Total'].sum():,.2f}")
        required_missing = []
        if not seller.strip():
            required_missing.append('Seller / company details')
        if not customer.strip():
            required_missing.append('Customer details')
        if not number.strip():
            required_missing.append('Quotation number')
        if valid_until < date:
            required_missing.append('Valid until must not precede quotation date')
        if required_missing:
            st.error('Complete required fields before exporting: ' + ', '.join(required_missing))
        else:
            st.success('Required quotation details complete. Exports are ready.')
        try:
            pdf = commercial_pdf(approved_lines, meta, currency, logo.getvalue() if logo else None)
            xlsx = commercial_excel(approved_lines, meta, currency)
            c1, c2 = st.columns(2)
            with c1:
                st.download_button('Download Professional Quotation (PDF)', pdf,
                                   file_name='QuoteFast_v093_Quotation.pdf', mime='application/pdf', disabled=bool(required_missing))
            with c2:
                st.download_button('Download Professional Quotation (Excel)', xlsx,
                                   file_name='QuoteFast_v093_Quotation.xlsx',
                                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', disabled=bool(required_missing))
        except ImportError as exc:
            st.error(f'Missing Python dependency: {exc}. Add reportlab to requirements.txt.')
        except Exception as exc:
            st.error(f'Could not create quotation: {exc}')

st.divider()
st.caption('QuoteFast v0.9.3 — rule-based prototype. General mode uses exact-description matching, not engineering verification. Human approval required.')
