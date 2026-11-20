#!/usr/bin/env python3
"""네이버 메일(IMAP)에서 SCM 재고 메일의 최신 엑셀 첨부를 내려받는다.

환경변수: NAVER_ID, NAVER_PW (네이버 '애플리케이션 비밀번호' 권장)
사용: python3 fetch_stock.py --out inputs/stock.xlsx
"""
import argparse, email, imaplib, os, sys
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime

SUBJECT_KEY = "아케이 품목별 재고"


def dec(s):
    try:
        return str(make_header(decode_header(s or "")))
    except Exception:
        return s or ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="inputs/stock.xlsx")
    ap.add_argument("--scan", type=int, default=80, help="최근 몇 통까지 볼지")
    a = ap.parse_args()
    uid, pw = os.environ.get("NAVER_ID"), os.environ.get("NAVER_PW")
    if not uid or not pw:
        sys.exit("NAVER_ID / NAVER_PW 가 필요해요")

    M = imaplib.IMAP4_SSL("imap.naver.com", 993)
    M.login(uid, pw)
    M.select("INBOX", readonly=True)
    _, data = M.search(None, "ALL")
    ids = data[0].split()[-a.scan:]
    best = None
    for i in reversed(ids):
        _, h = M.fetch(i, "(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])")
        msg = email.message_from_bytes(h[0][1])
        subj = dec(msg.get("Subject"))
        if SUBJECT_KEY in subj.replace("  ", " "):
            best = (i, subj, msg.get("Date"))
            break
    if not best:
        sys.exit(f"최근 {a.scan}통에서 '{SUBJECT_KEY}' 메일을 찾지 못했어요")

    _, raw = M.fetch(best[0], "(RFC822)")
    msg = email.message_from_bytes(raw[0][1])
    for part in msg.walk():
        fn = dec(part.get_filename())
        if fn and fn.lower().endswith((".xlsx", ".xls", ".xlsm")):
            os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
            open(a.out, "wb").write(part.get_payload(decode=True))
            when = parsedate_to_datetime(best[2]).astimezone().strftime("%m.%d %H:%M") if best[2] else ""
            label = f"{best[1]} ({when} 수신)"
            open(a.out + ".label", "w", encoding="utf-8").write(label)
            print("저장:", a.out, "|", label, "|", fn)
            break
    else:
        sys.exit("메일에 엑셀 첨부가 없어요: " + best[1])
    M.logout()


if __name__ == "__main__":
    main()
