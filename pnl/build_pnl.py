#!/usr/bin/env python3
"""올리브영 손익 구글시트(엑셀로 내려받은 것) → 대표님용 손익 대시보드 데이터(pnl/data.enc)

사용: DASH_PW=비번 python3 pnl/build_pnl.py 올리브영손익.xlsx [--provisional 가매출.json]
"""
import argparse, base64, datetime as dt, json, os, sys
from collections import defaultdict
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE); sys.path.insert(0, os.path.dirname(HERE))
from parse_sheet import load
from build import encrypt, ITER, KST, kdf

MAIN = "조선미녀"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--provisional", help="자동 수집한 가매출 행(JSON 배열)")
    ap.add_argument("--out", default=HERE)
    a = ap.parse_args()
    pw = os.environ["DASH_PW"]
    today = dt.datetime.now(KST).date().isoformat()

    rows = load(a.xlsx)
    for r in rows:
        # 시트에 수기로 넣은 행은 실매출. 단, 이번 달 안에서 입고일이 아직 안 온 금액 행만 가매출로 본다
        # (25.12 탭처럼 연도를 잘못 적은 날짜는 월과 안 맞아서 가매출로 잡지 않음)
        fut = r.get("inb") and r["inb"] >= today and r["inb"][:7] == r["month"] and (r["salesT"] or r["margin"])
        r["status"] = "가" if fut else "실"
    if a.provisional and os.path.exists(a.provisional):
        for r in json.load(open(a.provisional, encoding="utf-8")):
            r["status"] = r.get("status", "가")
            rows.append(r)

    months = defaultdict(lambda: defaultdict(float))
    for r in rows:
        m = months[r["month"]]
        for k in ("salesT", "costT", "edi", "logi", "margin"):
            m[k] += r[k]
            if r["brand"] == MAIN:
                m[k + "_J"] += r[k]
            if r["status"] == "가":
                m[k + "_P"] += r[k]
    mlist = [{"m": k, **{kk: round(v) for kk, v in sorted(months[k].items())}} for k in sorted(months)]

    data = {
        "built": dt.datetime.now(KST).strftime("%Y-%m-%d %H:%M"),
        "source": os.path.basename(a.xlsx),
        "months": mlist,
        "rows": [{k: (round(v) if isinstance(v, float) else v) for k, v in r.items()} for r in rows],
    }
    raw = json.dumps(data, ensure_ascii=False).encode()
    meta_path = os.path.join(a.out, "meta.json")
    salt = os.urandom(16)
    if os.path.exists(meta_path):
        salt = base64.b64decode(json.load(open(meta_path))["salt"])
    open(os.path.join(a.out, "data.enc"), "wb").write(encrypt(pw, salt, raw))
    json.dump({"salt": base64.b64encode(salt).decode(), "iter": ITER, "built": data["built"]}, open(meta_path, "w"))
    last = mlist[-1]
    print(f"{len(rows)}행 · {mlist[0]['m']}~{last['m']} · {last['m']} 매출 {last.get('salesT',0):,} 마진 {last.get('margin',0):,}")


if __name__ == "__main__":
    main()
