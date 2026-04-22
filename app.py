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
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from extractor import extract_invoice
from sheets import append_to_sheet

# ── Logging ───────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s'
)
log = logging.getLogger(__name__)

# ── App ───────────────────────────────────────────────────────────
app = Flask(__name__)

# ✅ CORS محدد — مش مفتوح على كل حاجة
ALLOWED_ORIGINS = os.environ.get(
    'ALLOWED_ORIGINS',
    'http://localhost:5000,http://127.0.0.1:5000'
).split(',')

CORS(app, origins=ALLOWED_ORIGINS)

# ✅ حد حجم الطلب — 50MB كحد أقصى للـ batch
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

MAX_FILE_SIZE = 20 * 1024 * 1024  # 20MB per file

# ✅ Credentials من البيئة فقط — مش من المستخدم
SA_JSON = os.environ.get('SA_JSON', '')


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


def _process_single(file, threshold):
    """معالجة ملف واحد — تُستخدم في الـ batch."""
    pdf_bytes, err = _validate_pdf(file)
    if err:
        return {'ok': False, 'filename': file.filename, 'error': err}
    try:
        data = extract_invoice(pdf_bytes, file.filename, threshold)
        return {'ok': True, 'data': data}
    except Exception as e:
        log.error(f'خطأ في معالجة {file.filename}: {e}')
        return {'ok': False, 'filename': file.filename, 'error': str(e)}


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

    threshold = int(request.form.get('threshold', 50))

    try:
        data = extract_invoice(pdf_bytes, request.files['file'].filename, threshold)
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

    threshold = int(request.form.get('threshold', 50))
    results   = [None] * len(files)

    # ✅ معالجة متوازية — كل الملفات في نفس الوقت
    with ThreadPoolExecutor(max_workers=4) as executor:
        future_to_index = {
            executor.submit(_process_single, f, threshold): i
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


@app.route('/api/send-to-sheets', methods=['POST'])
def send_to_sheets():
    """إرسال النتائج لـ Google Sheets."""
    body = request.get_json()
    if not body:
        return jsonify({'ok': False, 'error': 'لا يوجد بيانات'}), 400

    rows       = body.get('rows', [])
    sheet_id   = body.get('sheetId', '').strip()
    sheet_name = body.get('sheetName', 'فواتير الموردين').strip()

    # ✅ تحقق من المدخلات
    if not rows:
        return jsonify({'ok': False, 'error': 'لا يوجد صفوف للإرسال'}), 400
    if not sheet_id:
        return jsonify({'ok': False, 'error': 'لم يتم تحديد Sheet ID'}), 400

    # ✅ Credentials من السيرفر فقط
    if not SA_JSON:
        return jsonify({
            'ok': False,
            'error': 'لم يتم ضبط SA_JSON على السيرفر — راجع ملف .env'
        }), 500

    # ✅ تحقق إن الـ SA_JSON صالح
    try:
        json.loads(SA_JSON)
    except json.JSONDecodeError:
        return jsonify({'ok': False, 'error': 'SA_JSON غير صالح'}), 500

    try:
        added = append_to_sheet(rows, sheet_id, sheet_name, SA_JSON)
        log.info(f'Sheets: أضاف {added} صف إلى {sheet_id}')
        return jsonify({'ok': True, 'added': added})
    except Exception as e:
        log.error(f'خطأ في Sheets: {e}')
        return jsonify({'ok': False, 'error': str(e)}), 500


@app.route('/api/health')
def health():
    """فحص حالة السيرفر."""
    return jsonify({
        'ok':      True,
        'status':  'running',
        'version': '5.0',
        'sheets':  bool(SA_JSON),   # هل Sheets جاهز؟
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
    port  = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
    app.run(host='0.0.0.0', port=port, debug=debug)
