#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extractor.py — Invoice data extraction (QR → Text → OCR)
"""

import re
import base64
import io
from datetime import datetime
import fitz
from PIL import Image
import pytesseract
import numpy as np

try:
    from pyzbar.pyzbar import decode as pyzbar_decode
    PYZBAR_OK = True
except ImportError:
    PYZBAR_OK = False

DEFAULT_KW = {
    'invoice_num': [
        'رقم الفاتورة', 'فاتورة رقم', 'رقم المستند', 'رقم الضريبة', 'Invoice No',
        'Invoice Number', 'Tax Invoice No', 'Bill No', 'Receipt No', 'INV', 'FAT'
    ],
    'supplier': [
        'اسم المورد', 'المورد', 'البائع', 'اسم البائع', 'شركة', 'مؤسسة', 'Vendor',
        'Supplier', 'Seller', 'Issued By', 'From', 'Merchant', 'Company'
    ],
    'vat': [
        'الرقم الضريبي', 'رقم ضريبي', 'الرقم الضريبي للمورد', 'VAT No', 'VAT Number',
        'TRN', 'Tax No', 'Tax Number', 'VAT Registration No'
    ],
    'date': [
        'تاريخ الفاتورة', 'تاريخ الإصدار', 'تاريخ', 'التاريخ', 'Invoice Date',
        'Issue Date', 'Date', 'Bill Date'
    ],
    'desc': [
        'البيان', 'الوصف', 'وصف', 'الخدمة', 'الصنف', 'المنتج', 'Description',
        'Details', 'Service', 'Item', 'Product', 'Particulars'
    ],
    'subtotal': [
        'المجموع قبل الضريبة', 'الإجمالي قبل الضريبة', 'إجمالي قبل الضريبة',
        'المبلغ قبل الضريبة', 'قبل الضريبة', 'صافي القيمة', 'وعاء الضريبة',
        'المبلغ الخاضع للضريبة', 'Taxable Amount', 'Subtotal', 'Sub Total',
        'Net Amount', 'Amount Before VAT', 'before VAT'
    ],
    'vat_amt': [
        'مبلغ الضريبة', 'ضريبة القيمة المضافة', 'قيمة الضريبة', 'الضريبة',
        'VAT Amount', 'VAT', 'Tax Amount', 'Value Added Tax'
    ],
    'total': [
        'الإجمالي بعد الضريبة', 'إجمالي شامل الضريبة', 'الإجمالي شامل الضريبة',
        'الإجمالي النهائي', 'المبلغ الإجمالي', 'الإجمالي', 'إجمالي الفاتورة',
        'Grand Total', 'Total Amount', 'Total Incl VAT', 'Amount Due',
        'Net Payable', 'after VAT'
    ],
}

KW = DEFAULT_KW

NOISE_WORDS = (
    'tax invoice', 'فاتورة ضريبية', 'simplified tax invoice', 'invoice',
    'فاتورة', 'commercial registration', 'سجل تجاري', 'page ', 'صفحة'
)


def _merge_kw(custom_kw=None):
    merged = {k: list(v) for k, v in DEFAULT_KW.items()}
    if not isinstance(custom_kw, dict):
        return merged
    for key, cfg in custom_kw.items():
        words = cfg.get('words') if isinstance(cfg, dict) else cfg
        if key not in merged or not isinstance(words, list):
            continue
        for word in words:
            word = str(word).strip()
            if word and word not in merged[key]:
                merged[key].append(word)
    return merged


def _normalize_text(text):
    text = (text or '').replace('\u200f', ' ').replace('\u200e', ' ')
    text = text.replace('\xa0', ' ').replace('٫', '.').replace('٬', ',')
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def _clean_value(value):
    value = re.sub(r'\s+', ' ', str(value or '')).strip(' :-#|،,')
    value = re.sub(r'\b(SAR|ر\.?س\.?|ريال سعودي|ريال)\b$', '', value, flags=re.IGNORECASE).strip()
    return value


def _to_float(s):
    try:
        raw = str(s)
        raw = raw.translate(str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789'))
        raw = raw.replace('٫', '.').replace('٬', ',')
        m = re.search(r'-?\d[\d,،\s]*(?:\.\d+)?', raw)
        if not m:
            return None
        n = float(re.sub(r'[,،\s]', '', m.group(0)))
        return n if n > 0 else None
    except Exception:
        return None


def _amount_candidates(text):
    candidates = []
    amount_re = r'((?:SAR|ر\.?س\.?|ريال)?\s*\d[\d,،\s]{0,15}(?:\.\d{1,3})?\s*(?:SAR|ر\.?س\.?|ريال)?)'
    for line in _normalize_text(text).splitlines():
        clean = _clean_value(line)
        if not clean:
            continue
        for match in re.finditer(amount_re, clean, re.IGNORECASE):
            amount = _to_float(match.group(1))
            if amount:
                candidates.append((amount, clean))
    return candidates


def _keyword_score(line, words):
    line_l = line.lower()
    return sum(1 for w in words if str(w).lower() in line_l)


def _find_amount(text, field, kw):
    words = kw.get(field, [])
    best = None
    for amount, line in _amount_candidates(text):
        score = _keyword_score(line, words)
        if field == 'vat_amt' and re.search(r'\b15\s*%|١٥\s*%', line):
            score += 1
        if field == 'vat_amt' and re.search(r'شامل|بعد|الإجمالي|اجمالي|grand|total|due|payable|incl', line, re.IGNORECASE):
            score -= 3
        if field == 'vat_amt' and re.search(r'قبل|صافي|وعاء|خاضع|before|subtotal|sub total|net amount|taxable', line, re.IGNORECASE):
            score -= 3
        if field == 'vat_amt' and re.search(r'رقم|number|no\.?|trn|registration', line, re.IGNORECASE):
            score -= 3
        if field == 'total' and re.search(r'قبل|before|subtotal|taxable', line, re.IGNORECASE):
            score -= 2
        if field == 'subtotal' and re.search(r'بعد|شامل|grand|due|payable|total amount', line, re.IGNORECASE):
            score -= 2
        if score <= 0:
            continue
        if not best or score > best[0] or (score == best[0] and amount > best[1]):
            best = (score, amount)
    if best:
        return best[1]
    return None



def _find_text_field(text, field, kw, max_len=100):
    for w in kw.get(field, []):
        pat = r'(?:' + re.escape(w) + r')[\s:：#\-|]*([^\n\r]{2,' + str(max_len) + r'})'
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            value = _clean_value(m.group(1))
            if value and not _looks_like_label(value):
                return value[:max_len]
    return ''


def _looks_like_label(value):
    value_l = value.lower()
    return any(w in value_l for w in NOISE_WORDS) and len(value) < 25


def _find_invoice_number(text, kw):
    for w in kw.get('invoice_num', []):
        pat = r'(?:' + re.escape(w) + r')[\s:.#\-|]*(?:no\.?|number|رقم)?[\s:.#\-|]*([A-Z0-9][A-Z0-9\-\/_.]{1,40})'
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            value = _clean_value(m.group(1))
            if not re.fullmatch(r'\d{10,15}', value):
                return value[:40]
    patterns = [
        r'\b(INV[-\s]?\d{2,}[\w\-\/]*)\b',
        r'\b(FAT[-\s]?\d{2,}[\w\-\/]*)\b',
        r'\b(TI[-\s]?\d{2,}[\w\-\/]*)\b',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return _clean_value(m.group(1))[:40]
    return ''


def _find_vat_number(text, kw):
    value = _find_text_field(text, 'vat', kw, 50)
    m = re.search(r'\b3\d{13}\b', value or text)
    if m:
        return m.group(0)
    m = re.search(r'\b\d{15}\b', value or text)
    return m.group(0) if m else _clean_value(value)[:50]


def _normalize_date(value):
    value = _clean_value(value)
    value = value.translate(str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789'))
    patterns = [
        (r'\b(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\b', ('y', 'm', 'd')),
        (r'\b(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})\b', ('d', 'm', 'y')),
    ]
    for pat, order in patterns:
        m = re.search(pat, value)
        if not m:
            continue
        parts = dict(zip(order, m.groups()))
        y = int(parts['y'])
        if y < 100:
            y += 2000
        try:
            return datetime(y, int(parts['m']), int(parts['d'])).strftime('%Y-%m-%d')
        except ValueError:
            continue
    return value[:30]


def _find_date(text, kw):
    value = _find_text_field(text, 'date', kw, 60)
    if value:
        normalized = _normalize_date(value)
        if normalized:
            return normalized
    for pat in [
        r'\b\d{4}[-/.]\d{1,2}[-/.]\d{1,2}\b',
        r'\b\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\b',
    ]:
        m = re.search(pat, text)
        if m:
            return _normalize_date(m.group(0))
    return ''


def _find_supplier(text, kw):
    supplier = _find_text_field(text, 'supplier', kw, 100)
    if supplier:
        return supplier

    lines = [_clean_value(l) for l in text.splitlines()]
    candidates = []
    for idx, line in enumerate(lines[:25]):
        if not (4 <= len(line) <= 100):
            continue
        low = line.lower()
        if any(w in low for w in NOISE_WORDS):
            continue
        if re.search(r'\b\d{4}[-/]\d{1,2}\b|\b3\d{13}\b|@|www\.|https?://', line):
            continue
        if re.fullmatch(r'[\d\s\-\/\\.:،,.%]+', line):
            continue
        score = 0
        if re.search(r'شركة|مؤسسة|est\.?|co\.?|company|trading|مطعم|مكتب|مصنع', line, re.IGNORECASE):
            score += 3
        if re.search(r'[\u0600-\u06ffA-Za-z]', line):
            score += 1
        score -= idx * 0.05
        candidates.append((score, line))
    if not candidates:
        return ''
    return sorted(candidates, reverse=True)[0][1][:80]


def _find_description(text, kw):
    desc = _find_text_field(text, 'desc', kw, 140)
    if desc:
        return desc[:120]

    lines = [_clean_value(l) for l in text.splitlines()]
    for line in lines:
        if not (8 <= len(line) <= 140):
            continue
        if re.search(r'\d+\.\d{2}|%|\b3\d{13}\b|فاتورة|invoice|total|vat|ضريبة|الإجمالي', line, re.IGNORECASE):
            continue
        if re.search(r'[\u0600-\u06ffA-Za-z]', line):
            return line[:120]
    return ''


def _infer_amounts(text, subtotal, vat_amt, total):
    amounts = sorted({a for a, _ in _amount_candidates(text) if 0 < a < 10**9})

    if total and vat_amt and not subtotal and total > vat_amt:
        subtotal = round(total - vat_amt, 2)
    if subtotal and vat_amt and not total:
        total = round(subtotal + vat_amt, 2)
    if total and subtotal and not vat_amt and total > subtotal:
        vat_amt = round(total - subtotal, 2)

    if not total and amounts:
        total = max(amounts)
    if total and not vat_amt:
        for a in amounts:
            if abs(a - (total * 0.15 / 1.15)) < max(1, total * 0.01):
                vat_amt = a
                break
    if subtotal and not vat_amt:
        vat_amt = round(subtotal * 0.15, 2)
    if total and not subtotal and vat_amt and total > vat_amt:
        subtotal = round(total - vat_amt, 2)
    if not subtotal and total:
        subtotal = round(total / 1.15, 2)
        if not vat_amt:
            vat_amt = round(total - subtotal, 2)

    return subtotal, vat_amt, total

def parse_zatca_qr(raw):
    try:
        raw += '=' * (-len(raw) % 4)
        data = base64.b64decode(raw)
        fields = {}
        i = 0
        while i < len(data):
            if i + 1 >= len(data):
                break
            tag = data[i]
            i += 1
            length = data[i]
            i += 1
            if i + length > len(data):
                break
            value = data[i:i + length].decode('utf-8', errors='ignore')
            fields[tag] = value
            i += length

        if not fields.get(1) and not fields.get(2):
            return None

        total = _to_float(fields.get(4, ''))
        vat_amt = _to_float(fields.get(5, ''))
        subtotal = round(total - vat_amt, 2) if total and vat_amt else None

        return {
            'source': 'QR',
            'supplier': fields.get(1, '').strip(),
            'vat': fields.get(2, '').strip(),
            'date': fields.get(3, '')[:10],
            'invoice_num': '',
            'desc': '',
            'subtotal': '{:.2f}'.format(subtotal) if subtotal else '',
            'vat_amt': '{:.2f}'.format(vat_amt) if vat_amt else '',
            'total': '{:.2f}'.format(total) if total else '',
            'currency': 'SAR',
            'filename': '',
            'complete': bool(fields.get(1) and fields.get(2) and total),
        }
    except Exception:
        return None

def scan_qr_from_pdf(pdf_bytes):
    if not PYZBAR_OK:
        return None
    try:
        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        for page_num in range(min(3, len(doc))):
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
            img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
            for code in pyzbar_decode(img):
                raw = code.data.decode('utf-8', errors='ignore')
                result = parse_zatca_qr(raw)
                if result:
                    return result
    except Exception:
        pass
    return None

def scan_qr_from_text(pdf_bytes):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        for page in doc:
            text = page.get_text("rawtext")
            matches = re.findall(r'[A-Za-z0-9+/]{50,}={0,2}', text)
            for m in matches:
                result = parse_zatca_qr(m)
                if result:
                    return result
    except Exception:
        pass
    return None

def extract_text_from_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    return _normalize_text('\n'.join(page.get_text('text', sort=True) for page in doc))

def is_scanned(text, threshold=50):
    compact = re.sub(r'\s+', '', text or '')
    return len(compact) < threshold

def ocr_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    full = ''
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), colorspace=fitz.csGRAY, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes('png')))
        arr = np.array(img)
        thresh = ((arr > 155) * 255).astype(np.uint8)
        full += pytesseract.image_to_string(
            Image.fromarray(thresh),
            lang='ara+eng',
            config='--oem 3 --psm 6 -c preserve_interword_spaces=1'
        ) + '\n'
    return _normalize_text(full)

def parse_invoice_text(text, filename, source='نص', custom_kw=None):
    text = _normalize_text(text)
    kw = _merge_kw(custom_kw)

    invoice_num = _find_invoice_number(text, kw)
    supplier = _find_supplier(text, kw)
    vat = _find_vat_number(text, kw)
    date = _find_date(text, kw)
    desc = _find_description(text, kw)
    cm = re.search(r'\b(SAR|USD|EUR|AED|GBP|ريال)\b', text, re.IGNORECASE)
    currency = 'SAR' if not cm or 'ريال' in cm.group(0) else cm.group(0).upper()

    s = _find_amount(text, 'subtotal', kw)
    v = _find_amount(text, 'vat_amt', kw)
    t = _find_amount(text, 'total', kw)
    s, v, t = _infer_amounts(text, s, v, t)

    return {
        'source': source,
        'invoice_num': invoice_num[:40],
        'supplier': supplier[:80],
        'vat': vat,
        'date': date,
        'desc': desc[:120],
        'subtotal': '{:.2f}'.format(s) if s else '',
        'vat_amt': '{:.2f}'.format(v) if v else '',
        'total': '{:.2f}'.format(t) if t else '',
        'currency': currency[:5],
        'filename': filename,
        'complete': bool(supplier and vat and date and t),
    }

def _is_garbled(text):
    if not text or len(text.strip()) < 20:
        return True
    weird = sum(
        1 for c in text
        if ('\u0370' <= c <= '\u03ff') or ('\u0200' <= c <= '\u02ff') or ('\u0080' <= c <= '\u009f')
    )
    return (weird / max(len(text), 1)) > 0.15

def _merge_missing(primary, secondary):
    for key in ['invoice_num', 'supplier', 'vat', 'date', 'desc', 'subtotal', 'vat_amt', 'total', 'currency']:
        if not primary.get(key) and secondary.get(key):
            primary[key] = secondary[key]
    primary['complete'] = bool(primary.get('supplier') and primary.get('vat') and primary.get('date') and primary.get('total'))
    return primary


def extract_invoice(pdf_bytes, filename, ocr_threshold=50, custom_kw=None):
    qr = scan_qr_from_pdf(pdf_bytes)
    if not qr:
        qr = scan_qr_from_text(pdf_bytes)

    text = extract_text_from_pdf(pdf_bytes)

    if qr:
        qr['filename'] = filename
        if not _is_garbled(text):
            text_data = parse_invoice_text(text, filename, custom_kw=custom_kw)
            return _merge_missing(qr, text_data)
        return _merge_missing(qr, {'filename': filename})

    if is_scanned(text, ocr_threshold) or _is_garbled(text):
        try:
            ocr_text = ocr_pdf(pdf_bytes)
            return parse_invoice_text(ocr_text, filename, source='OCR', custom_kw=custom_kw)
        except Exception:
            return parse_invoice_text('', filename, source='OCR (فشل)', custom_kw=custom_kw)

    return parse_invoice_text(text, filename, source='نص', custom_kw=custom_kw)
