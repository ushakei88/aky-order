#!/usr/bin/env python3
"""회송 목록: 올리브원 납품확인서(inputs/r1) 검수 확정 수량 + CJ 납품완료보고서 수기 사유(inputs/returns.json.enc)"""
import json, os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "pnl")); sys.path.insert(0, HERE)


def build_returns(pw, r1_glob=os.path.join(HERE, "inputs", "r1", "*.json.enc"), reasons_path=os.path.join(HERE, "inputs", "returns.json.enc")):
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

