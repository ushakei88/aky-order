#!/usr/bin/env python3
"""올리브원 납품확인서(R1) 데이터 → 손익 시트 자동 탭 갱신

- 자동_품목  : 품목 한 줄씩 (작업시트와 같은 계산)
- 자동_월집계: 월 탭과 같은 모양 (납품건 × 브랜드 × 센터)
수량 기준: 납품수량(PDA 검수 확정, supQty) > 0 → 실매출 / 아니면 납품예정수량(supTobeQty) > 0 → 가매출 / 둘 다 없으면 넣지 않음

사용 (깃허브 Actions):
  GCP_SA_KEY=... DASH_PW=... python3 pnl/oy_sheet.py inputs/r1/*.json.enc [--dry]
"""
import argparse, datetime as dt, glob, json, os, re, sys, tempfile
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

SHEET_ID = "1wDfEEEL1Bl8Kr5Cg807E7KBM8PSsESMGnQfxaT6vgas"
KST = dt.timezone(dt.timedelta(hours=9))
EDI_RATE, LOGI_RATE = 0.04, 0.04
BRAND_FIX = {"하우스오브허": "하오허", "그래니": "그래니 샐러드"}
TAB_ITEM, TAB_SUM = "자동_품목", "자동_월집계"
ITEM_HEAD = ["상태", "발주일", "입고일", "센터", "센터코드", "전표", "상품코드", "상품명", "브랜드", "발주수량", "납품예정수량",
             "납품수량", "계산수량", "원단가", "올영매출(공급가)", "부가세율", "매입가(LAW)", "매입금액(공급가)", "EDI", "물류비", "이익(공급가)", "비고", "갱신시각"]
SUM_HEAD = ["상태", "구분", "발주일", "입고일", "브랜드", "센터", "아케이매입 공급가액", "세액", "합계금액", "올리브영매입 공급가액", "세액", "합계금액",
            "EDI 공급가액", "세액", "합계금액", "물류대행 공급가액", "세액", "합계금액", "마진", "마진율", "갱신시각"]


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def code(v):
    s = re.sub(r"\.0$", "", str(v or "").strip())
    return re.sub(r"\D", "", s)


def ymd(s):
    s = str(s or "")
    return f"{s[:4]}-{s[4:6]}-{s[6:8]}" if re.fullmatch(r"\d{8}", s) else s


def brand_of(name):
    b = (name or "").split()[0] if (name or "").split() else ""
    return BRAND_FIX.get(b, b)


def load_r1(paths, pw):
    from vault import key  # noqa
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    rows = []
    for p in paths:
        if p.endswith(".enc"):
            data = open(p, "rb").read()
            raw = AESGCM(key(pw, data[:16])).decrypt(data[16:28], data[28:], None)
            d = json.loads(raw)
        else:
            d = json.load(open(p, encoding="utf-8"))
        for r in d["rows"]:
            rows.append(dict(zip(d["cols"], r)))
    # 같은 전표·상품·센터는 마지막 파일 기준
    uniq = {}
    for g in rows:
        uniq[(g.get("cntrPoSlipNo"), g.get("cntrPoSlipSeq"), g.get("gdsCd"), g.get("stkinCntrCd"))] = g
    return list(uniq.values())


def calc(r1, law):
    now = dt.datetime.now(KST).strftime("%Y-%m-%d %H:%M")
    items, missing = [], set()
    for g in r1:
        sup, tobe = num(g.get("supQty")), num(g.get("supTobeQty"))
        if sup > 0:
            status, qty = "실매출", sup
        elif tobe > 0:
            status, qty = "가매출", tobe
        else:
            continue
        c = code(g.get("gdsCd"))
        center = "온라인" if g.get("stkinCntrCd") == "LA04" else "오프라인"
        price = num(g.get("costUprc"))
        vat = num(g.get("vatRt")) / 100 if g.get("vatRt") not in (None, "") else 0.1
        sales = price * qty
        cost_u = law.get(c)
        note = ""
        if cost_u is None:
            missing.add((c, g.get("gdsNm")))
            note = "LAW에 매입가 없음"
        cost = (cost_u or 0) * qty
        edi = sales * EDI_RATE
        logi = sales * LOGI_RATE if center == "오프라인" else 0
        items.append({
            "상태": status, "발주일": ymd(g.get("poYmd")), "입고일": ymd(g.get("stkinYmd")), "센터": center,
            "센터코드": g.get("stkinCntrCd"), "전표": g.get("cntrPoSlipNo"), "상품코드": c, "상품명": g.get("gdsNm"),
            "브랜드": brand_of(g.get("gdsNm")), "발주수량": num(g.get("poQty")), "납품예정수량": tobe, "납품수량": sup,
            "계산수량": qty, "원단가": price, "올영매출(공급가)": sales, "부가세율": vat, "매입가(LAW)": cost_u if cost_u is not None else "",
            "매입금액(공급가)": cost, "EDI": edi, "물류비": logi, "이익(공급가)": sales - cost - edi - logi, "비고": note, "갱신시각": now,
        })
    return items, sorted(missing)


def summarize(items):
    now = dt.datetime.now(KST).strftime("%Y-%m-%d %H:%M")
    grp = {}
    for it in items:
        k = (it["발주일"], it["입고일"], it["브랜드"], it["센터"])
        g = grp.setdefault(k, {"st": set(), "cost": 0, "costV": 0, "sales": 0, "salesV": 0, "edi": 0, "logi": 0})
        g["st"].add(it["상태"])
        g["cost"] += it["매입금액(공급가)"]; g["costV"] += it["매입금액(공급가)"] * it["부가세율"]
        g["sales"] += it["올영매출(공급가)"]; g["salesV"] += it["올영매출(공급가)"] * it["부가세율"]
        g["edi"] += it["EDI"]; g["logi"] += it["물류비"]
    out = []
    for (po, inb, brand, center), g in sorted(grp.items(), key=lambda x: (x[0][1], x[0][0], x[0][2], x[0][3])):
        costT, salesT = g["cost"] + g["costV"], g["sales"] + g["salesV"]
        ediT, logiT = g["edi"] * 1.1, g["logi"] * 1.1
        margin = salesT - costT - ediT - logiT
        out.append({"상태": "가매출" if "가매출" in g["st"] else "실매출", "구분": "정기", "발주일": po, "입고일": inb, "브랜드": brand, "센터": center,
                    "a": [g["cost"], g["costV"], costT, g["sales"], g["salesV"], salesT, g["edi"], g["edi"] * .1, ediT, g["logi"], g["logi"] * .1, logiT,
                          margin, (margin / costT) if costT else 0], "갱신시각": now})
    return out


def gs_client():
    import gspread
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_info(json.loads(os.environ["GCP_SA_KEY"]), scopes=[
        "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.readonly"])
    return gspread.authorize(creds), creds


def read_law(sh):
    vals = sh.worksheet("LAW").get_all_values()
    law = {}
    for r in vals[1:]:
        if len(r) > 3 and r[0] and r[3]:
            try:
                law[code(r[0].replace(",", ""))] = float(r[3].replace(",", ""))
            except ValueError:
                pass
    return law


def ensure_tab(sh, title, head, index):
    import gspread
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=2000, cols=len(head), index=index)
    return ws


def write_tabs(sh, items, sums, missing):
    rnd = lambda v: round(v) if isinstance(v, float) and abs(v) >= 1 else (round(v, 4) if isinstance(v, float) else v)
    ws = ensure_tab(sh, TAB_ITEM, ITEM_HEAD, 1)
    old = ws.get_all_values()
    keep = [r for r in old[1:] if r and r[2] and r[2] not in {i["입고일"] for i in items}]  # 이번에 다시 계산한 입고일은 새로 씀
    body = keep + [[rnd(it[h]) for h in ITEM_HEAD] for it in items]
    body.sort(key=lambda r: (r[2], r[3], r[8], r[7]))
    ws.clear()
    ws.update([ITEM_HEAD] + body, value_input_option="USER_ENTERED")
    ws.freeze(rows=1)

    # 월집계는 자동_품목 전체로 다시 만든다
    allitems = []
    for r in body:
        d = dict(zip(ITEM_HEAD, r))
        for k in ("올영매출(공급가)", "매입금액(공급가)", "EDI", "물류비", "부가세율"):
            d[k] = num(str(d[k]).replace(",", ""))
        allitems.append(d)
    sums = summarize(allitems)
    ws2 = ensure_tab(sh, TAB_SUM, SUM_HEAD, 2)
    ws2.clear()
    ws2.update([SUM_HEAD] + [[s["상태"], s["구분"], s["발주일"], s["입고일"], s["브랜드"], s["센터"], *[rnd(v) for v in s["a"]], s["갱신시각"]] for s in sums],
               value_input_option="USER_ENTERED")
    ws2.freeze(rows=1)
    return sums


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--dry", action="store_true", help="시트에 쓰지 않고 계산만")
    ap.add_argument("--law-xlsx", help="(--dry용) LAW를 엑셀에서 읽기")
    a = ap.parse_args()
    paths = sorted({p for f in a.files for p in glob.glob(f)})
    r1 = load_r1(paths, os.environ.get("DASH_PW", ""))
    if a.dry:
        import openpyxl
        wb = openpyxl.load_workbook(a.law_xlsx, data_only=True, read_only=True)
        law = {code(r[0]): float(r[3]) for r in list(wb["LAW"].iter_rows(values_only=True))[1:] if r[0] and isinstance(r[3], (int, float))}
        items, missing = calc(r1, law)
        sums = summarize(items)
        for s in sums:
            print(s["상태"], s["입고일"], s["브랜드"], s["센터"], "매출", round(s["a"][5]), "마진", round(s["a"][12]))
        print("LAW 없음:", missing)
        return
    gc, _ = gs_client()
    sh = gc.open_by_key(SHEET_ID)
    items, missing = calc(r1, read_law(sh))
    sums = write_tabs(sh, items, None, missing)
    print(f"자동_품목 {len(items)}줄 갱신 · 자동_월집계 {len(sums)}줄")
    if missing:
        print("LAW에 매입가 없음:", "; ".join(f"{c} {n}" for c, n in missing))
        open(os.environ.get("GITHUB_STEP_SUMMARY", os.devnull), "a").write("LAW에 매입가 없는 품목: " + ", ".join(n or c for c, n in missing) + "\n")


if __name__ == "__main__":
    main()
