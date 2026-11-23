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


def load_auto(path, manual):
    """시트의 '자동_월집계' 탭 → 손익 행. 수기 월 탭에 이미 있는 입고일은 건너뜀(중복 방지)"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "자동_월집계" not in wb.sheetnames:
        return []
    have = {(r["month"], r["inb"]) for r in manual}
    out = []
    for v in list(wb["자동_월집계"].iter_rows(values_only=True))[1:]:
        if not v or not v[3]:
            continue
        inb = str(v[3])[:10]
        if (inb[:7], inb) in have:
            continue
        f = lambda x: float(str(x).replace(",", "")) if x not in (None, "") else 0.0
        out.append({"month": inb[:7], "kind": v[1] or "정기", "po": str(v[2])[:10], "inb": inb, "brand": v[4], "center": v[5],
                    "cost": f(v[6]), "costT": f(v[8]), "sales": f(v[9]), "salesT": f(v[11]), "edi": f(v[14]), "logi": f(v[17]),
                    "margin": f(v[18]), "post": 0.0, "status": "가" if v[0] == "가매출" else "실", "auto": True})
    return out


def build_returns(r1_glob, reasons_path, pw):
    """회송 = 검수 확정 후 (납품예정수량 − 납품수량). 사유는 CJ 납품완료보고서 수기 메모를 옮겨 적은 표에서."""
    import glob as _g
    from oy_sheet import load_r1, code as _code, brand_of
    from vault import key
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    files = sorted(_g.glob(r1_glob))
    rows = load_r1(files, pw) if files else []
    price = {}
    for g in rows:
        if g.get("costUprc"):
            price[_code(g.get("gdsCd"))] = (float(g["costUprc"]), 0.0 if g.get("vatSpCd") == "M" else 0.1)
    reasons = []
    if reasons_path and os.path.exists(reasons_path):
        d = open(reasons_path, "rb").read()
        reasons = json.loads(AESGCM(key(pw, d[:16])).decrypt(d[16:28], d[28:], None))
    rk = {(r["date"], r["center"], r["code"]): r for r in reasons}
    out = {}
    for g in rows:
        sup, tobe = float(g.get("supQty") or 0), float(g.get("supTobeQty") or 0)
        done = g.get("cnfmYn") == "Y" and str(g.get("stkinStatCd")) == "4"
        if not done or tobe <= sup:
            continue
        s = str(g.get("stkinYmd")); date = f"{s[:4]}-{s[4:6]}-{s[6:8]}"
        c = _code(g.get("gdsCd")); cc = g.get("stkinCntrCd")
        unit = float(g.get("inptUnitPerQty") or 0)
        k = (date, cc, c)
        o = out.setdefault(k, {"date": date, "center": "온라인" if cc == "LA04" else "오프라인", "code": c, "name": g.get("gdsNm"),
                               "brand": brand_of(g.get("gdsNm")), "tried": 0, "delivered": 0, "unit": unit})
        o["tried"] += tobe; o["delivered"] += sup
    for k, r in rk.items():
        if k not in out:
            out[k] = {"date": r["date"], "center": "온라인" if r["center"] == "LA04" else "오프라인", "code": r["code"], "name": r["name"],
                      "brand": brand_of(r["name"]), "tried": r["qty"], "delivered": 0, "unit": r["qty"] / max(1, r.get("boxes", 1))}
    res = []
    for k, o in out.items():
        r = rk.get(k, {})
        o["returned"] = o["tried"] - o["delivered"]
        if not o["unit"] and r.get("boxes"):
            o["unit"] = r["qty"] / r["boxes"]
        o["cartons"] = round(o["returned"] / o["unit"], 1) if o["unit"] else r.get("boxes")
        p, v = price.get(o["code"], (0, 0.1))
        o["lost"] = round(o["returned"] * p * (1 + v))
        o["reason"] = r.get("reason", "사유 확인 전")
        o["memo"] = r.get("memo", "")
        o["source"] = r.get("source", "")
        res.append(o)
    res.sort(key=lambda x: (x["date"], x["center"], x["name"]), reverse=True)
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--provisional", help="자동 수집한 가매출 행(JSON 배열)")
    ap.add_argument("--out", default=HERE)
    ap.add_argument("--r1", default=os.path.join(os.path.dirname(HERE), "inputs", "r1", "*.json.enc"))
    ap.add_argument("--reasons", default=os.path.join(os.path.dirname(HERE), "inputs", "returns.json.enc"))
    a = ap.parse_args()
    pw = os.environ["DASH_PW"]
    today = dt.datetime.now(KST).date().isoformat()

    rows = load(a.xlsx)
    for r in rows:
        if r.get("status") in ("가", "실"):
            continue
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

    returns = build_returns(a.r1, a.reasons, pw)

    data = {
        "returns": returns,
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
