#!/usr/bin/env python3
"""서비스 계정으로 손익 시트를 엑셀로 내려받기: python3 pnl/export_sheet.py /tmp/oy_pnl.xlsx"""
import json, os, sys
import google.auth.transport.requests as gr
import requests
from google.oauth2.service_account import Credentials
SHEET_ID = "1wDfEEEL1Bl8Kr5Cg807E7KBM8PSsESMGnQfxaT6vgas"
creds = Credentials.from_service_account_info(json.loads(os.environ["GCP_SA_KEY"]), scopes=["https://www.googleapis.com/auth/drive.readonly"])
creds.refresh(gr.Request())
r = requests.get(f"https://www.googleapis.com/drive/v3/files/{SHEET_ID}/export",
                 params={"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                 headers={"Authorization": f"Bearer {creds.token}"}, timeout=120)
r.raise_for_status()
open(sys.argv[1], "wb").write(r.content)
print("내려받음", len(r.content), "bytes")
