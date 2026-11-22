#!/usr/bin/env python3
"""서비스 계정으로 손익 시트에 쓸 수 있는지 확인: '자동입력_연결확인' 탭에 한 줄 기록"""
import json, os, datetime as dt
import gspread
from google.oauth2.service_account import Credentials
SHEET_ID = "1wDfEEEL1Bl8Kr5Cg807E7KBM8PSsESMGnQfxaT6vgas"
creds = Credentials.from_service_account_info(json.loads(os.environ["GCP_SA_KEY"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"])
sh = gspread.authorize(creds).open_by_key(SHEET_ID)
name = "자동입력_연결확인"
try:
    ws = sh.worksheet(name)
except gspread.WorksheetNotFound:
    ws = sh.add_worksheet(title=name, rows=20, cols=4)
    ws.append_row(["기록 시각(한국)", "내용"])
now = dt.datetime.now(dt.timezone(dt.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M")
ws.append_row([now, "aky-sheet-bot 연결 확인 — 이 탭은 지워도 돼요"])
print("OK", sh.title, [w.title for w in sh.worksheets()][:5])
