#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py — Flask Invoice Extractor API
"""

import os
import json
from flask import Flask, request, jsonify, render_template, send_from_directory
from flask_cors import CORS
from extractor import extract_invoice
from sheets import append_to_sheet

app = Flask(__name__)
CORS(app)

MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 MB

# ── Routes ────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/extract', methods=['POST'])
def extract():
    """Extract invoice data from uploaded PDF."""
    if 'file' not in request.files:
        return jsonify({'ok': False, 'error': 'لم يتم إرسال أي ملف'}), 400

    file = request.files['file']
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({'ok': False, 'error': 'الملف ليس PDF'}), 400

    pdf_bytes = file.read()
    if len(pdf_bytes) > MAX_FILE_SIZE:
        return jsonify({'ok': False, 'error': 'حجم الملف أكبر من 20MB'}), 400

    threshold = int(request.form.get('threshold', 50))

    try:
        data = extract_invoice(pdf_bytes, file.filename, threshold)
        return jsonify({'ok': True, 'data': data})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@app.route('/api/extract-batch', methods=['POST'])
def extract_batch():
    """Extract from multiple PDFs at once."""
    files = request.files.getlist('files')
    if not files:
        return jsonify({'ok': False, 'error': 'لم يتم إرسال أي ملفات'}), 400

    threshold = int(request.form.get('threshold', 50))
    results = []

    for file in files:
        if not file.filename.lower().endswith('.pdf'):
            results.append({'ok': False, 'filename': file.filename, 'error': 'ليس PDF'})
            continue
        pdf_bytes = file.read()
        if len(pdf_bytes) > MAX_FILE_SIZE:
            results.append({'ok': False, 'filename': file.filename, 'error': 'أكبر من 20MB'})
            continue
        try:
            data = extract_invoice(pdf_bytes, file.filename, threshold)
            results.append({'ok': True, 'data': data})
        except Exception as e:
            results.append({'ok': False, 'filename': file.filename, 'error': str(e)})

    return jsonify({'ok': True, 'results': results, 'total': len(results)})


@app.route('/api/send-to-sheets', methods=['POST'])
def send_to_sheets():
    """Append invoice rows to Google Sheets."""
    body = request.get_json()
    if not body:
        return jsonify({'ok': False, 'error': 'لا يوجد بيانات'}), 400

    rows         = body.get('rows', [])
    sheet_id     = body.get('sheetId', '')
    sheet_name   = body.get('sheetName', 'فواتير الموردين')
    credentials  = body.get('credentials', '')

    if not rows:
        return jsonify({'ok': False, 'error': 'لا يوجد صفوف'}), 400
    if not sheet_id:
        return jsonify({'ok': False, 'error': 'لم يتم تحديد Sheet ID'}), 400
    if not credentials:
        return jsonify({'ok': False, 'error': 'لم يتم تحديد بيانات الاعتماد'}), 400

    try:
        added = append_to_sheet(rows, sheet_id, sheet_name, credentials)
        return jsonify({'ok': True, 'added': added})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@app.route('/api/health')
def health():
    return jsonify({'ok': True, 'status': 'running', 'version': '4.0'})


# ── Run ───────────────────────────────────────────────────────────
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
    app.run(host='0.0.0.0', port=port, debug=debug)
