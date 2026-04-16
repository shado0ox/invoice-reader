#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sheets.py — Google Sheets integration via Service Account
"""

import json
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

HEADERS = [
    '#', 'المصدر', 'رقم الفاتورة', 'اسم المورد', 'الرقم الضريبي',
    'التاريخ', 'البيان', 'قبل الضريبة', 'الضريبة', 'بعد الضريبة',
    'العملة', 'اسم الملف', 'تاريخ الإدخال', 'الحالة'
]

def append_to_sheet(rows: list, sheet_id: str, sheet_name: str, credentials_json: str) -> int:
    creds_dict = json.loads(credentials_json) if isinstance(credentials_json, str) else credentials_json
    creds      = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    service    = build('sheets', 'v4', credentials=creds)
    sheet      = service.spreadsheets()
    now_str    = datetime.now().strftime('%Y-%m-%d %H:%M')

    # Check if sheet exists, create if not
    meta = sheet.get(spreadsheetId=sheet_id).execute()
    sheet_titles = [s['properties']['title'] for s in meta['sheets']]

    if sheet_name not in sheet_titles:
        body = {'requests': [{'addSheet': {'properties': {'title': sheet_name}}}]}
        sheet.batchUpdate(spreadsheetId=sheet_id, body=body).execute()
        # Add header row
        sheet.values().update(
            spreadsheetId=sheet_id,
            range=f'{sheet_name}!A1',
            valueInputOption='RAW',
            body={'values': [HEADERS]}
        ).execute()
        _format_header(service, sheet_id, sheet_name)

    # Get current row count
    result = sheet.values().get(
        spreadsheetId=sheet_id,
        range=f'{sheet_name}!A:A'
    ).execute()
    current_rows = len(result.get('values', []))

    # Prepare rows
    values = []
    for i, r in enumerate(rows):
        row_num = current_rows + i + 1
        values.append([
            row_num - 1,          # # (exclude header)
            r.get('source',''),
            r.get('invoice_num',''),
            r.get('supplier',''),
            r.get('vat',''),
            r.get('date',''),
            r.get('desc',''),
            r.get('subtotal',''),
            r.get('vat_amt',''),
            r.get('total',''),
            r.get('currency','SAR'),
            r.get('filename',''),
            now_str,
            'مكتملة' if r.get('complete') else 'تحقق يدوي',
        ])

    if values:
        sheet.values().append(
            spreadsheetId=sheet_id,
            range=f'{sheet_name}!A1',
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body={'values': values}
        ).execute()

    return len(values)


def _format_header(service, sheet_id: str, sheet_name: str):
    """Format header row with dark blue background."""
    meta = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sid  = next(s['properties']['sheetId'] for s in meta['sheets']
                if s['properties']['title'] == sheet_name)
    requests = [{
        'repeatCell': {
            'range': {'sheetId': sid, 'startRowIndex': 0, 'endRowIndex': 1},
            'cell': {
                'userEnteredFormat': {
                    'backgroundColor': {'red': 0.1, 'green': 0.16, 'blue': 0.37},
                    'textFormat': {'foregroundColor': {'red':1,'green':1,'blue':1},
                                   'bold': True, 'fontSize': 10},
                    'horizontalAlignment': 'RIGHT',
                    'verticalAlignment': 'MIDDLE',
                }
            },
            'fields': 'userEnteredFormat'
        }
    }, {
        'updateSheetProperties': {
            'properties': {'sheetId': sid, 'gridProperties': {'frozenRowCount': 1},
                           'rightToLeft': True},
            'fields': 'gridProperties.frozenRowCount,rightToLeft'
        }
    }]
    service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={'requests': requests}
    ).execute()
