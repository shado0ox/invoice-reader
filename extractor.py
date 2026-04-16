#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extractor.py — Invoice data extraction engine
Supports: ZATCA QR Code · Text PDF · Scanned PDF (OCR)
"""

import re
import base64
import struct
from datetime import datetime
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import cv2
import numpy as np
import io

# ── ZATCA QR (TLV / Base64) ───────────────────────────────────────
def parse_zatca_qr(raw: str) -> dict | None:
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
        raw_date = fields.get(3, '')[:10]
        return {
            'source':    'QR',
            'supplier':  fields.get(1, '').strip(),
            'vat':       fields.get(2, '').strip(),
            'date':      raw_date,
            'invoice_num': '',
            'desc':      '',
            'subtotal':  f'{subtotal:.2f}' if subtotal else '',
            'vat_amt':   f'{vat_amt:.2f}'  if vat_amt  else '',
            'total':     f'{total:.2f}'    if total     else '',
            'currency':  'SAR',
            'complete':  bool(fields.get(1) and fields.get(2) and total),
        }
    except Exception:
        return None


def scan_qr_from_pdf(pdf_bytes: bytes) -> dict | None:
    """Render each page at high-res and scan for QR codes."""
    try:
        from pyzbar.pyzbar import decode as pyzbar_decode
        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        for page_num in range(min(3, len(doc))):
            page = doc[page_num]
            mat  = fitz.Matrix(3, 3)          # 3x scale ≈ 216 DPI
            pix  = page.get_pixmap(matrix=mat)
            img  = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
            codes = pyzbar_decode(img)
            for code in codes:
                raw = code.data.decode('utf-8', errors='ignore')
                result = parse_zatca_qr(raw)
                if result:
                    return result
    except Exception:
        pass
    return None


# ── PDF text extraction ───────────────────────────────────────────
def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    doc  = fitz.open(stream=pdf_bytes, filetype='pdf')
    text = ''
    for page in doc:
        text += page.get_text() + '
'
    return text


def is_scanned(text: str, threshold: int = 50) -> bool:
    return len(text.replace(' ', '').replace('
', '')) < threshold


# ── OCR on scanned PDF ────────────────────────────────────────────
def ocr_pdf(pdf_bytes: bytes) -> str:
    doc  = fitz.open(stream=pdf_bytes, filetype='pdf')
    full = ''
    for page in doc:
        mat = fitz.Matrix(2.5, 2.5)
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
        img_bytes = pix.tobytes('png')
        img = Image.open(io.BytesIO(img_bytes))
        # Preprocessing
        img_np  = np.array(img)
        _, thresh = cv2.threshold(img_np, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        img_clean = Image.fromarray(thresh)
        text = pytesseract.image_to_string(
            img_clean,
            lang='ara+eng',
            config='--oem 3 --psm 3'
        )
        full += text + '
'
    return full


# ── Text parser ───────────────────────────────────────────────────
KW = {
    'invoice_num': ['رقم الفاتورة','فاتورة رقم','Invoice No','Invoice Number','INV','FAT'],
    'supplier':    ['المورد','Supplier','Vendor','Issued By','مقدم من','البائع','شركة'],
    'vat':         ['رقم ضريبي','الرقم الضريبي','VAT No','VAT Number','TRN','Tax No'],
    'date':        ['التاريخ','تاريخ الفاتورة','Invoice Date','Date'],
    'desc':        ['البيان','الوصف','Description','Details','Service','Item','بند'],
    'subtotal':    ['المجموع قبل','إجمالي قبل','الإجمالي ريال','المجموع','Subtotal',
                    'Net Amount','before VAT','قبل الضريبة','صافي القيمة','وعاء الضريبة','Taxable Amount'],
    'vat_amt':     ['مبلغ الضريبة','ضريبة القيمة المضافة','VAT Amount','Tax Amount','الضريبة'],
    'total':       ['الإجمالي بعد','إجمالي شامل','الإجمالي النهائي','Grand Total',
                    'Total Amount','after VAT','المبلغ الإجمالي'],
}

def _to_float(s: str) -> float | None:
    try:
        n = float(re.sub(r'[,،\s]', '', str(s)))
        return n if n > 0 else None
    except Exception:
        return None

def _find_amount(text: str, field: str) -> float | None:
    for w in KW.get(field, []):
        esc = re.escape(w)
        m = re.search(rf'(?:{esc})[\s\S]{{0,15}}?([\d,،.]+)', text, re.IGNORECASE)
        if m:
            n = _to_float(m.group(1))
            if n: return n
    return None

def _find_text_field(text: str, field: str) -> str:
    for w in KW.get(field, []):
        esc = re.escape(w)
        m = re.search(rf'(?:{esc})[:\s]+([^\n\r،,]{{3,80}})', text, re.IGNORECASE)
        if m:
            return re.sub(r'\s+', ' ', m.group(1)).strip()
    return ''

def parse_invoice_text(text: str, filename: str, source: str = 'نص') -> dict:
    # Invoice number
    invoice_num = ''
    for w in KW['invoice_num']:
        m = re.search(rf'(?:{re.escape(w)})[\s:.#-]*(\w[\w\-\/]{{1,30}})', text, re.IGNORECASE)
        if m: invoice_num = m.group(1).strip(); break

    # Supplier
    supplier = _find_text_field(text, 'supplier')
    if not supplier:
        lines = [l.strip() for l in text.split('\n') if 3 < len(l.strip()) < 80 and not l.strip()[0].isdigit()]
        supplier = lines[0] if lines else ''

    # VAT number
    vat = _find_text_field(text, 'vat')
    if not vat:
        m = re.search(r'\b3[0-9]{13}\b', text)
        if m: vat = m.group(0)

    # Date
    date = _find_text_field(text, 'date')
    if not date:
        for pat in [r'\b(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})\b', r'\b(\d{4}-\d{2}-\d{2})\b']:
            m = re.search(pat, text)
            if m: date = m.group(1); break

    desc     = _find_text_field(text, 'desc')
    currency = (re.search(r'\b(SAR|USD|EUR|AED|GBP|ريال)\b', text, re.IGNORECASE) or [None])[0] or 'SAR'

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
        'subtotal':    f'{s:.2f}' if s else '',
        'vat_amt':     f'{v:.2f}' if v else '',
        'total':       f'{t:.2f}' if t else '',
        'currency':    currency[:5],
        'filename':    filename,
        'complete':    bool(supplier and date and t),
    }


# ── Main pipeline ─────────────────────────────────────────────────
def extract_invoice(pdf_bytes: bytes, filename: str, ocr_threshold: int = 50) -> dict:
    """
    Pipeline:
    1. Try ZATCA QR Code
    2. Try text extraction
    3. If scanned → OCR
    """
    # Step 1: QR
    qr_data = scan_qr_from_pdf(pdf_bytes)
    if qr_data:
        qr_data['filename'] = filename
        return qr_data

    # Step 2: Text
    text = extract_text_from_pdf(pdf_bytes)
    if not is_scanned(text, ocr_threshold):
        result = parse_invoice_text(text, filename, source='نص')
        return result

    # Step 3: OCR
    try:
        ocr_text = ocr_pdf(pdf_bytes)
        result   = parse_invoice_text(ocr_text, filename, source='OCR')
        return result
    except Exception as e:
        return parse_invoice_text('', filename, source='OCR (فشل)')
