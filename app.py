#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py — Flask Invoice Extractor API v5
"""
from dotenv import load_dotenv  # ← أضف
load_dotenv()                   # ← أضف

import os
import json
import logging
import zipfile
from io import BytesIO
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
from extractor import extract_invoice

# ── Logging ───────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s'
)
log = logging.getLogger(__name__)

# ── App ───────────────────────────────────────────────────────────
app = Flask(__name__)

# الواجهة قد تُفتح من ملف محلي أو من بورت مختلف، لذلك نفتح CORS لمسارات API فقط.
CORS(app, resources={r"/api/*": {"origins": "*"}})

# ✅ حد حجم الطلب — 50MB كحد أقصى للـ batch
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

MAX_FILE_SIZE = 20 * 1024 * 1024  # 20MB per file

# ── Helpers ───────────────────────────────────────────────────────

def _validate_pdf(file):
    """تحقق من الملف قبل المعالجة."""
    if not file or not file.filename:
        return None, 'لم يتم إرسال أي ملف'
    if not file.filename.lower().endswith('.pdf'):
        return None, 'الملف ليس PDF'
    pdf_bytes = file.read()
    if len(pdf_bytes) > MAX_FILE_SIZE:
        return None, 'حجم الملف أكبر من 20MB'
    if len(pdf_bytes) < 100:
        return None, 'الملف فارغ أو تالف'
    return pdf_bytes, None


def _parse_custom_kw(raw):
    """اقرأ كلمات البحث المخصصة القادمة من الواجهة."""
    if not raw:
        return None
    try:
        parsed = json.loads(raw)
        return parsed if isinstance(parsed, dict) else None
    except json.JSONDecodeError:
        return None


def _safe_int(value, default=50):
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _process_single(file, threshold, custom_kw=None):
    """معالجة ملف واحد — تُستخدم في الـ batch."""
    pdf_bytes, err = _validate_pdf(file)
    if err:
        return {'ok': False, 'filename': file.filename, 'error': err}
    try:
        data = extract_invoice(pdf_bytes, file.filename, threshold, custom_kw)
        return {'ok': True, 'data': data}
    except Exception as e:
        log.error(f'خطأ في معالجة {file.filename}: {e}')
        return {'ok': False, 'filename': file.filename, 'error': str(e)}


EXPORT_HEADERS = [
    '#', 'المصدر', 'رقم الفاتورة', 'اسم المورد', 'الرقم الضريبي',
    'التاريخ', 'البيان', 'قبل الضريبة', 'الضريبة', 'بعد الضريبة',
    'العملة', 'اسم الملف', 'تاريخ التصدير', 'الحالة'
]


def _row_from_invoice(idx, row, exported_at):
    return [
        idx,
        row.get('source', ''),
        row.get('invoice_num', ''),
        row.get('supplier', ''),
        row.get('vat', ''),
        row.get('date', ''),
        row.get('desc', ''),
        row.get('subtotal', ''),
        row.get('vat_amt', ''),
        row.get('total', ''),
        row.get('currency', 'SAR'),
        row.get('filename', ''),
        exported_at,
        'مكتملة' if row.get('complete') else 'تحقق يدوي',
    ]


def _xlsx_col_name(index):
    name = ''
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def _xml_escape(value):
    return (
        str(value if value is not None else '')
        .replace('&', '&amp;')
        .replace('<', '&lt;')
        .replace('>', '&gt;')
        .replace('"', '&quot;')
    )


def _cell_xml(row_num, col_num, value, header=False):
    ref = f'{_xlsx_col_name(col_num)}{row_num}'
    style = ' s="1"' if header else ''
    if col_num in {1, 8, 9, 10} and not header:
        try:
            number = float(str(value).replace(',', ''))
            return f'<c r="{ref}"{style}><v>{number}</v></c>'
        except (TypeError, ValueError):
            pass
    return f'<c r="{ref}" t="inlineStr"{style}><is><t>{_xml_escape(value)}</t></is></c>'


def _build_xlsx(rows):
    exported_at = datetime.now().strftime('%Y-%m-%d %H:%M')
    values = [EXPORT_HEADERS]
    values.extend(_row_from_invoice(i, row, exported_at) for i, row in enumerate(rows, start=1))

    sheet_rows = []
    for r_idx, row in enumerate(values, start=1):
        cells = ''.join(_cell_xml(r_idx, c_idx, value, header=(r_idx == 1))
                        for c_idx, value in enumerate(row, start=1))
        sheet_rows.append(f'<row r="{r_idx}">{cells}</row>')

    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="1" width="6" customWidth="1"/>
    <col min="2" max="6" width="18" customWidth="1"/>
    <col min="7" max="7" width="34" customWidth="1"/>
    <col min="8" max="10" width="15" customWidth="1"/>
    <col min="11" max="14" width="18" customWidth="1"/>
  </cols>
  <sheetData>{''.join(sheet_rows)}</sheetData>
  <autoFilter ref="A1:N{len(values)}"/>
</worksheet>'''

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="فواتير الموردين" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Arial"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1A2A5E"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="right"/></xf>
  </cellXfs>
</styleSheet>'''

    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    workbook_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''

    out = BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types_xml)
        zf.writestr('_rels/.rels', rels_xml)
        zf.writestr('xl/workbook.xml', workbook_xml)
        zf.writestr('xl/_rels/workbook.xml.rels', workbook_rels_xml)
        zf.writestr('xl/styles.xml', styles_xml)
        zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)
    out.seek(0)
    return out


# ── Routes ────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/extract', methods=['POST'])
def extract():
    """استخراج بيانات فاتورة واحدة."""
    if 'file' not in request.files:
        return jsonify({'ok': False, 'error': 'لم يتم إرسال أي ملف'}), 400

    pdf_bytes, err = _validate_pdf(request.files['file'])
    if err:
        return jsonify({'ok': False, 'error': err}), 400

    threshold = _safe_int(request.form.get('threshold'), 50)
    custom_kw = _parse_custom_kw(request.form.get('kw', ''))

    try:
        data = extract_invoice(pdf_bytes, request.files['file'].filename, threshold, custom_kw)
        log.info(f"✅ استخرج: {data.get('filename')} [{data.get('source')}]")
        return jsonify({'ok': True, 'data': data})
    except Exception as e:
        log.error(f'خطأ في الاستخراج: {e}')
        return jsonify({'ok': False, 'error': str(e)}), 500


@app.route('/api/extract-batch', methods=['POST'])
def extract_batch():
    """استخراج من أكثر من ملف PDF في نفس الوقت."""
    files = request.files.getlist('files')
    if not files:
        return jsonify({'ok': False, 'error': 'لم يتم إرسال أي ملفات'}), 400

    # ✅ حد أقصى 20 ملف في نفس الوقت
    if len(files) > 20:
        return jsonify({'ok': False, 'error': 'الحد الأقصى 20 ملف في كل طلب'}), 400

    threshold = _safe_int(request.form.get('threshold'), 50)
    custom_kw = _parse_custom_kw(request.form.get('kw', ''))
    results   = [None] * len(files)

    # ✅ معالجة متوازية — كل الملفات في نفس الوقت
    with ThreadPoolExecutor(max_workers=4) as executor:
        future_to_index = {
            executor.submit(_process_single, f, threshold, custom_kw): i
            for i, f in enumerate(files)
        }
        for future in as_completed(future_to_index):
            i = future_to_index[future]
            try:
                results[i] = future.result()
            except Exception as e:
                results[i] = {
                    'ok': False,
                    'filename': files[i].filename,
                    'error': str(e)
                }

    success = sum(1 for r in results if r and r.get('ok'))
    log.info(f'Batch: {success}/{len(files)} نجح')
    return jsonify({'ok': True, 'results': results, 'total': len(results)})


@app.route('/api/export-xlsx', methods=['POST'])
def export_xlsx():
    """تصدير النتائج إلى ملف Excel محلي بدون Google Sheets."""
    body = request.get_json()
    if not body:
        return jsonify({'ok': False, 'error': 'لا يوجد بيانات'}), 400

    rows = body.get('rows', [])
    if not rows:
        return jsonify({'ok': False, 'error': 'لا يوجد صفوف للتصدير'}), 400

    output = _build_xlsx(rows)
    filename = f"invoices_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/api/send-to-sheets', methods=['POST'])
def send_to_sheets():
    """المشروع لم يعد يعتمد على Google Sheets."""
    return jsonify({
        'ok': False,
        'error': 'تم إلغاء الاعتماد على Google Sheets. استخدم تصدير Excel بدلاً منه.'
    }), 410


@app.route('/api/health')
def health():
    """فحص حالة السيرفر."""
    return jsonify({
        'ok':      True,
        'status':  'running',
        'version': '6.0',
        'sheets':  False,
        'exports': ['xlsx', 'csv'],
        'ocr':     _check_tesseract()
    })


# ── Internal Checks ───────────────────────────────────────────────

def _check_tesseract():
    """تحقق إن Tesseract متثبت."""
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


# ── Error Handlers ────────────────────────────────────────────────

@app.errorhandler(413)
def too_large(e):
    return jsonify({'ok': False, 'error': 'حجم الطلب أكبر من 50MB'}), 413

@app.errorhandler(404)
def not_found(e):
    return jsonify({'ok': False, 'error': 'المسار غير موجود'}), 404

@app.errorhandler(500)
def server_error(e):
    return jsonify({'ok': False, 'error': 'خطأ داخلي في السيرفر'}), 500


# ── Run ───────────────────────────────────────────────────────────

if __name__ == '__main__':
    port  = int(os.environ.get('PORT', 5050))
    debug = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
    app.run(host='0.0.0.0', port=port, debug=debug)
