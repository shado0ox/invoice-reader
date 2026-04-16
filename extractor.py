#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import re, base64, io
import fitz
from PIL import Image
import pytesseract
import cv2
import numpy as np

def parse_zatca_qr(raw):
    try:
        data = base64.b64decode(raw)
        fields = {}
        i = 0
        while i < len(data):
            tag = data[i]; i += 1
            length = data[i]; i += 1
            value = data[i:i+length].decode('utf-8', errors='ignore')
            fields[tag] = value
            i += length
        if not fields.get(1) and not fields.get(2):
            return None
        total   = _to_float(fields.get(4, ''))
        vat_amt = _to_float(fields.get(5, ''))
        subtotal = round(total - vat_amt, 2) if total and vat_amt else None
        return {
            'source':      'QR',
            'supplier':    fields.get(1, '').strip(),
            'vat':         fields.get(2, '').strip(),
            'date':        fields.get(3, '')[:10],
            'invoice_num': '',
            'desc':        '',
            'subtotal':    '{:.2f}'.format(subtotal) if subtotal else '',
            'vat_amt':     '{:.2f}'.format(vat_amt)  if vat_amt  else '',
            'total':       '{:.2f}'.format(total)    if total    else '',
            'currency':    'SAR',
            'filename':    '',
            'complete':    bool(fields.get(1) and fields.get(2) and total),
        }
    except Exception:
        return None

def scan_qr_from_pdf(pdf_bytes):
    try:
        from pyzbar.pyzbar import decode as pyzbar_decode
        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        for page_num in range(min(3, len(doc))):
            page = doc[page_num]
            pix  = page.get_pixmap(matrix=fitz.Matrix(3, 3))
            img  = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
            for code in pyzbar_decode(img):
                raw = code.data.decode('utf-8', errors='ignore')
                result = parse_zatca_qr(raw)
                if result:
                    return result
    except Exception:
        pass
    return None

def extract_text_from_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    return '\n'.join(page.get_text() for page in doc)

def is_scanned(text, threshold=50):
    return len(text.replace(' ', '').replace('\n', '')) < threshold

def ocr_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    full = ''
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), colorspace=fitz.csGRAY)
        img = Image.open(io.BytesIO(pix.tobytes('png')))
        arr = np.array(img)
        _, thresh = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        full += pytesseract.image_to_string(Image.fromarray(thresh), lang='ara+eng', config='--oem 3 --psm 3') + '\n'
    return full

KW = {
    'invoice_num': ['رقم الفاتورة','فاتورة رقم','Invoice No','Invoice Number','INV','FAT'],
    'supplier':    ['المورد','Supplier','Vendor','Issued By','مقدم من','البائع','شركة'],
    'vat':         ['رقم ضريبي','الرقم الضريبي','VAT No','VAT Number','TRN','Tax No'],
    'date':        ['التاريخ','تاريخ الفاتورة','Invoice Date','Date'],
    'desc':        ['البيان','الوصف','Description','Details','Service','Item','بند'],
    'subtotal':    ['المجموع قبل','إجمالي قبل','المجموع','Subtotal','Net Amount',
                    'before VAT','قبل الضريبة','صافي القيمة','وعاء الضريبة','Taxable Amount'],
    'vat_amt':     ['مبلغ الضريبة','ضريبة القيمة المضافة','VAT Amount','Tax Amount','الضريبة'],
    'total':       ['الإجمالي بعد','إجمالي شامل','الإجمالي النهائي','Grand Total',
                    'Total Amount','after VAT','المبلغ الإجمالي'],
}

def _to_float(s):
    try:
        n = float(re.sub(r'[,،\s]', '', str(s)))
        return n if n > 0 else None
    except Exception:
        return None

def _find_amount(text, field):
    for w in KW.get(field, []):
        m = re.search(r'(?:' + re.escape(w) + r')[\s\S]{0,15}?([\d,،.]+)', text, re.IGNORECASE)
        if m:
            n = _to_float(m.group(1))
            if n: return n
    return None

def _find_text_field(text, field):
    for w in KW.get(field, []):
        m = re.search(r'(?:' + re.escape(w) + r')[:\s]+([^\n\r،,]{3,80})', text, re.IGNORECASE)
        if m:
            return re.sub(r'\s+', ' ', m.group(1)).strip()
    return ''

def parse_invoice_text(text, filename, source='نص'):
    invoice_num = ''
    for w in KW['invoice_num']:
        m = re.search(r'(?:' + re.escape(w) + r')[\s:.#-]*(\w[\w\-\/]{1,30})', text, re.IGNORECASE)
        if m: invoice_num = m.group(1).strip(); break
    supplier = _find_text_field(text, 'supplier')
    if not supplier:
        lines = [l.strip() for l in text.split('\n') if 3 < len(l.strip()) < 80 and not l.strip()[:1].isdigit()]
        supplier = lines[0] if lines else ''
    vat = _find_text_field(text, 'vat')
    if not vat:
        m = re.search(r'\b3[0-9]{13}\b', text)
        if m: vat = m.group(0)
    date = _find_text_field(text, 'date')
    if not date:
        for pat in [r'\b(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})\b', r'\b(\d{4}-\d{2}-\d{2})\b']:
            m = re.search(pat, text)
            if m: date = m.group(1); break
    desc = _find_text_field(text, 'desc')
    cm = re.search(r'\b(SAR|USD|EUR|AED|GBP|ريال)\b', text, re.IGNORECASE)
    currency = cm.group(0) if cm else 'SAR'
    s = _find_amount(text, 'subtotal')
    v = _find_amount(text, 'vat_amt')
    t = _find_amount(text, 'total')
    if s and v and not t: t = round(s + v, 2)
    if t and v and not s: s = round(t - v, 2)
    if t and s and not v: v = round(t - s, 2)
    if not v and s:       v = round(s * 0.15, 2)
    if not t and s and v: t = round(s + v, 2)
    return {
        'source':      source,
        'invoice_num': invoice_num[:40],
        'supplier':    supplier[:80],
        'vat':         vat,
        'date':        date,
        'desc':        desc[:120],
        'subtotal':    '{:.2f}'.format(s) if s else '',
        'vat_amt':     '{:.2f}'.format(v) if v else '',
        'total':       '{:.2f}'.format(t) if t else '',
        'currency':    currency[:5],
        'filename':    filename,
        'complete':    bool(supplier and date and t),
    }

def extract_invoice(pdf_bytes, filename, ocr_threshold=50):
    qr = scan_qr_from_pdf(pdf_bytes)
    if qr:
        qr['filename'] = filename
        return qr
    text = extract_text_from_pdf(pdf_bytes)
    if not is_scanned(text, ocr_threshold):
        return parse_invoice_text(text, filename, source='نص')
    try:
        return parse_invoice_text(ocr_pdf(pdf_bytes), filename, source='OCR')
    except Exception:
        return parse_invoice_text('', filename, source='OCR (فشل)')