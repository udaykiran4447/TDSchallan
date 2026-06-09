"""GSTR-3B PDF parser — robust extractor for all sections."""

import re
import pdfplumber
from pathlib import Path


def _num(val):
    s = str(val).strip().replace(',', '').replace('L', '')
    if s in ('-', '', '--', 'NA', ' '):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _get4(pattern, text):
    m = re.search(pattern + r'\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)', text, re.DOTALL)
    return [_num(m.group(i)) for i in range(1, 5)] if m else [0.0]*4


def _parse_payment_rows(section_text):
    """Parse 4 payment rows (IGST/CGST/SGST/Cess) from a payment section block."""
    lines = [l for l in section_text.split('\n') if l.strip()]
    rows = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        row_type, rest = None, line
        for prefix, lbl in [('Integrated', 'IGST'), ('Central', 'CGST'),
                             ('State/UT', 'SGST'), ('Cess', 'Cess')]:
            if line.startswith(prefix):
                row_type = lbl
                rest = line[len(prefix):].strip()
                break
        if row_type:
            main_toks = re.findall(r'[\d.]+|-', rest)
            # Check continuation line (orphan digit fragments after label word)
            if i+1 < len(lines):
                nxt = lines[i+1].strip()
                if nxt.startswith('tax') or re.match(r'^[\d\s.\-]+$', nxt):
                    cont_rest = nxt[3:].strip() if nxt.startswith('tax') else nxt
                    cont_toks = re.findall(r'[\d.]+', cont_rest)
                    all_toks = main_toks + cont_toks
                    # Merge adjacent number fragments
                    merged = []
                    for t in all_toks:
                        if merged and t.startswith('.') and re.match(r'^\d+$', merged[-1]):
                            merged[-1] += t
                        elif merged and re.match(r'^\d+$', t) and re.match(r'^\d+\.\d?$', merged[-1]) and len(merged[-1].split('.')[-1]) < 2:
                            merged[-1] += t
                        else:
                            merged.append(t)
                    main_toks = merged
                    i += 2
                else:
                    i += 1
            else:
                i += 1
            vals = [_num(t) for t in main_toks[:10]]
            while len(vals) < 10:
                vals.append(0.0)
            rows.append((row_type, vals))
        else:
            i += 1
    return rows


def extract_gstr3b_data(pdf_path):
    """Extract all sections from a GSTR-3B PDF. Returns a dict."""
    full_text = ''
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + '\n'

    d = {}

    # ── Metadata ──────────────────────────────────────────────────────────────
    m = re.search(r'Period\s+(\w+)', full_text)
    d['period'] = m.group(1) if m else 'Unknown'
    m = re.search(r'Year\s+([\d\-]+)', full_text)
    d['year'] = m.group(1) if m else ''
    m = re.search(r'GSTIN of the supplier\s+(\S+)', full_text)
    d['gstin'] = m.group(1) if m else ''
    m = re.search(r'Legal name of the registered person\s+(.+)', full_text)
    d['legal_name'] = m.group(1).strip() if m else ''

    # ── 3.1 Supplies ─────────────────────────────────────────────────────────
    m = re.search(
        r'\(a\) Outward taxable supplies \(other than zero rated.*?(\d[\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)',
        full_text)
    d['3_1a'] = [_num(m.group(i)) for i in range(1, 6)] if m else [0.0]*5

    m = re.search(r'\(b\) Outward taxable supplies \(zero rated\)\s*([\d.]+)\s+([\d.]+)', full_text)
    d['3_1b'] = [_num(m.group(1)), _num(m.group(2)), 0, 0, 0] if m else [0.0]*5

    m = re.search(r'\(c\s*\)\s*Other outward supplies \(nil rated.*?\)\s*([\d.]+)', full_text)
    d['3_1c'] = [_num(m.group(1)), 0, 0, 0, 0] if m else [0.0]*5

    m = re.search(
        r'\(d\) Inward supplies \(liable to reverse charge\)\s*L?([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)',
        full_text)
    d['3_1d'] = [_num(m.group(i)) for i in range(1, 6)] if m else [0.0]*5

    m = re.search(r'\(e\) Non-GST outward supplies\s*([\d.]+)', full_text)
    d['3_1e'] = [_num(m.group(1)), 0, 0, 0, 0] if m else [0.0]*5

    # ── 3.1.1 ECO supplies ────────────────────────────────────────────────────
    m = re.search(
        r'\(i\) Taxable supplies on which electronic commerce operator pays.*?\]\s*([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)',
        full_text, re.DOTALL)
    d['3_1_1i'] = [_num(m.group(i)) for i in range(1, 6)] if m else [0.0]*5

    m = re.search(r'\(ii\) Taxable supplies made by registered person.*?operator\]\s*([\d.]+)', full_text, re.DOTALL)
    d['3_1_1ii'] = [_num(m.group(1)), 0, 0, 0, 0] if m else [0.0]*5

    # ── 3.2 Inter-state ───────────────────────────────────────────────────────
    m = re.search(r'Supplies made to Unregistered Persons\s+([\d.]+)\s+([\d.]+)', full_text)
    d['3_2_unreg'] = [_num(m.group(1)), _num(m.group(2))] if m else [0.0, 0.0]
    m = re.search(r'Supplies made to Composition Taxable\s*Persons?\s+([\d.]+)\s+([\d.]+)', full_text)
    d['3_2_comp'] = [_num(m.group(1)), _num(m.group(2))] if m else [0.0, 0.0]
    m = re.search(r'Supplies made to UIN holders\s+([\d.]+)\s+([\d.]+)', full_text)
    d['3_2_uin'] = [_num(m.group(1)), _num(m.group(2))] if m else [0.0, 0.0]

    # ── 4. ITC ────────────────────────────────────────────────────────────────
    d['itc_a1'] = _get4(r'\(1\) Import of goods', full_text)
    d['itc_a2'] = _get4(r'\(2\) Import of services', full_text)
    d['itc_a3'] = _get4(r'\(3\) Inward supplies liable to reverse charge \(other than 1 & 2 above\)', full_text)
    d['itc_a4'] = _get4(r'\(4\) Inward supplies from ISD', full_text)
    d['itc_a5'] = _get4(r'\(5\) All other ITC', full_text)
    d['itc_b1'] = _get4(r'\(1\) As per rules 38,42 & 43', full_text)
    d['itc_b2'] = _get4(r'\(2\) Others\s*\n', full_text)
    d['itc_c']  = _get4(r'C\.\s*Net ITC available \(A-B\)', full_text)
    d['itc_d1'] = _get4(r'\(1\) ITC reclaimed which was reversed', full_text)
    m = re.search(r'16\(4\).*?PoS rules\s+([\d.]+)\s+([\d.]+)\s+(\d[\d.]*)\s+([\d.]+)', full_text, re.DOTALL)
    if not m:
        m = re.search(r'16\(4\).*?PoS rules\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)', full_text, re.DOTALL)
    d['itc_d2'] = [_num(m.group(i)) for i in range(1, 5)] if m else [0.0]*4

    # ── 5. Exempt inward + 5.1 Interest ──────────────────────────────────────
    m = re.search(r'From a supplier under composition scheme.*?supply\s+([\d.]+)\s+([\d.]+)', full_text, re.DOTALL)
    d['sec5_comp'] = [_num(m.group(1)), _num(m.group(2))] if m else [0.0, 0.0]
    m = re.search(r'Non GST supply\s+([\d.]+)\s+([\d.]+)', full_text)
    d['sec5_nongst'] = [_num(m.group(1)), _num(m.group(2))] if m else [0.0, 0.0]
    m = re.search(r'Interest Paid\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)', full_text)
    d['interest_paid'] = [_num(m.group(i)) for i in range(1, 5)] if m else [0.0]*4
    m = re.search(r'Late fee\s+[-–]\s+([\d.]+)\s+([\d.]+)', full_text)
    d['late_fee'] = [0, _num(m.group(1)), _num(m.group(2)), 0] if m else [0.0]*4

    # ── 6.1 Payment of tax ────────────────────────────────────────────────────
    pay_full = re.search(r'6\.1 Payment of tax(.+?)Breakup of tax', full_text, re.DOTALL)
    if pay_full:
        ps = pay_full.group(1)
        a_block = re.search(r'\(A\) Other than reverse charge(.+?)\(B\) Reverse charge', ps, re.DOTALL)
        b_block = re.search(r'\(B\) Reverse charge.+?9\(5\)(.+?)$', ps, re.DOTALL)

        key_map = {'IGST': 0, 'CGST': 1, 'SGST': 2, 'Cess': 3}
        section_keys_A = ['pay_a_igst', 'pay_a_cgst', 'pay_a_sgst', 'pay_a_cess']
        section_keys_B = ['pay_b_igst', 'pay_b_cgst', 'pay_b_sgst', 'pay_b_cess']

        for block, keys in [(a_block, section_keys_A), (b_block, section_keys_B)]:
            # init with zeros
            for k in keys:
                d[k] = [0.0]*10
            if block:
                rows = _parse_payment_rows(block.group(1))
                for row_type, vals in rows:
                    idx = key_map.get(row_type)
                    if idx is not None and idx < len(keys):
                        d[keys[idx]] = vals
    else:
        for k in ['pay_a_igst','pay_a_cgst','pay_a_sgst','pay_a_cess',
                  'pay_b_igst','pay_b_cgst','pay_b_sgst','pay_b_cess']:
            d[k] = [0.0]*10

    return d


def month_order(month_str):
    order = ['april','may','june','july','august','september',
             'october','november','december','january','february','march']
    return order.index(month_str.lower()) if month_str.lower() in order else 99
