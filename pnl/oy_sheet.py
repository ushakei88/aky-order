#!/usr/bin/env python3
"""올리브원 납품확인서(R1) 데이터 → 손익 시트 자동 탭 갱신

- 자동_품목  : 품목 한 줄씩 (작업시트와 같은 계산)
- 자동_월집계: 월 탭과 같은 모양 (납품건 × 브랜드 × 센터)
수량 기준: 납품수량(PDA 검수 확정, supQty) > 0 → 실매출 / 아니면 납품예정수량(supTobeQty) > 0 → 가매출 / 둘 다 없으면 넣지 않음

사용 (깃허브 Actions):
  GCP_SA_KEY=... DASH_PW=... python3 pnl/oy_sheet.py inputs/r1/*.json.enc [--dry]
"""
import argparse, datetime as dt, glob, json, os, re, sys, tempfile
import gspread.utils  # noqa: E402  (Actions에서만 필요)
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
        done = g.get("cnfmYn") == "Y" and str(g.get("stkinStatCd")) == "4"  # 검수(PDA) 확정
        if done or sup > 0:
            status, qty = "실매출", sup
        elif tobe > 0:
            status, qty = "가매출", tobe
        else:
            continue
        if qty <= 0:
            continue  # 확정됐는데 0개 = 미납·회송
        c = code(g.get("gdsCd"))
        center = "온라인" if g.get("stkinCntrCd") == "LA04" else "오프라인"
        price = num(g.get("costUprc"))
        vat = 0.0 if g.get("vatSpCd") == "M" else (num(g.get("vatRt")) / 100 if g.get("vatRt") not in (None, "") else 0.1)  # M = 면세
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


def write_item_tab(sh, items):
    rnd = lambda v: round(v) if isinstance(v, float) and abs(v) >= 1 else (round(v, 4) if isinstance(v, float) else v)
    ws = ensure_tab(sh, TAB_ITEM, ITEM_HEAD, 1)
    old = ws.get_all_values()
    keep = [r for r in old[1:] if r and r[2] and r[2] not in {i["입고일"] for i in items}]
    body = keep + [[rnd(it[h]) for h in ITEM_HEAD] for it in items]
    body.sort(key=lambda r: (str(r[2]), str(r[3]), str(r[8]), str(r[7])))
    ws.clear()
    ws.update([ITEM_HEAD] + body, value_input_option="USER_ENTERED")
    ws.freeze(rows=1)


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


# ---------------------------------------------------------------- 월 탭(26.11 등)에 이어서 누적
AUTO_TAG_REAL = "자동·실매출"
AUTO_TAG_PROV = "가매출 · 전량 납품 가정, 납품 후 실매출로 수정 예정"


def month_tab_name(inb):
    return f"{inb[2:4]}.{inb[5:7]}"


def row_formulas(r):
    return [f'=if($D{r}="청담소녀",0,F{r}*0.1)', f"=F{r}+G{r}", None, f'=if($D{r}="청담소녀",0,I{r}*0.1)', f"=I{r}+J{r}",
            f"=I{r}*0.04", f"=L{r}*10%", f"=L{r}+M{r}", f'=IF(or($E{r}="온라인",$E{r}="글로벌"),0,I{r}*0.04)', f"=O{r}*10%",
            f"=O{r}+P{r}", f"=K{r}-H{r}-N{r}-Q{r}+T{r}", f"=R{r}/H{r}"]


def ensure_month_tab(sh, name):
    import gspread
    for w in sh.worksheets():
        if w.title.strip().lstrip("'") == name:
            return w, False
    # 직전 달 탭을 복사해서 새 달 탭을 만든다 (데이터·비용 금액은 비움)
    allws = sh.worksheets()
    tabs = [w for w in allws if re.fullmatch(r"\d\d\.\d\d", w.title.strip().lstrip("'"))]
    src = max(tabs, key=lambda w: w.title.strip().lstrip("'"))
    ws = sh.duplicate_sheet(src.id, insert_sheet_index=[w.id for w in allws].index(src.id), new_sheet_name=("'" if src.title.startswith("'") else "") + name)
    vals = ws.get_all_values()
    cost_hdr = next((i for i, r in enumerate(vals) if r and r[0].strip() == "<비용>"), len(vals))
    last = max((i for i in range(10, cost_hdr) if len(vals[i]) > 3 and vals[i][3]), default=10)
    if last >= 10:
        ws.batch_clear([f"A11:U{last + 1}"])
    if cost_hdr < len(vals):
        ws.batch_clear([f"C{cost_hdr + 3}:E{len(vals)}"])
    return ws, True


def write_month_tabs(sh, items):
    """자동_품목을 월 탭 형식(납품건 × 브랜드 × 센터)으로 묶어 월 탭 데이터 아래에 이어 쓴다.
    자동으로 쓴 줄은 U열에 표시가 있어서, 다음 실행 때 지우고 다시 쓴다(가매출 → 실매출 수정)."""
    by_tab = {}
    for it in items:
        by_tab.setdefault(month_tab_name(it["입고일"]), []).append(it)
    report = []
    for name, its in sorted(by_tab.items()):
        ws, created = ensure_month_tab(sh, name)
        vals = ws.get_all_values()
        cost_hdr = next((i for i, r in enumerate(vals) if r and r[0].strip() == "<비용>"), len(vals))
        # 기존 자동 줄 지우기 (아래에서부터)
        auto_rows = [i + 1 for i in range(10, cost_hdr) if len(vals[i]) > 20 and vals[i][20].startswith(("자동", "가매출"))]
        if auto_rows and max(auto_rows) - min(auto_rows) + 1 == len(auto_rows):
            ws.delete_rows(min(auto_rows), max(auto_rows))
        else:
            for r in sorted(auto_rows, reverse=True):
                ws.delete_rows(r)
        vals = ws.get_all_values()
        cost_hdr = next((i for i, r in enumerate(vals) if r and r[0].strip() == "<비용>"), len(vals))
        manual_last = max((i + 1 for i in range(10, cost_hdr) if len(vals[i]) > 3 and vals[i][3]), default=10)
        groups = {}
        for it in its:
            k = (it["발주일"], it["입고일"], it["브랜드"], it["센터"])
            g = groups.setdefault(k, {"cost": 0.0, "sales": 0.0, "prov": False})
            g["cost"] += num(it["매입금액(공급가)"]); g["sales"] += num(it["올영매출(공급가)"])
            g["prov"] |= it["상태"] == "가매출"
        keys = sorted(groups, key=lambda k: (k[1], k[0], k[2], k[3]))
        n = len(keys)
        if not n:
            continue
        free = cost_hdr - manual_last  # 비용 표 위 빈 줄 수
        need = n + 2 - free            # 비용 표 위로 빈 줄 2개는 남김
        start = manual_last + 1
        if need > 0:
            ws.insert_rows([[""] * 21] * need, row=start)
        rows, prev = [], None
        for i, k in enumerate(keys):
            r = start + i
            po, inb, brand, center = k
            g = groups[k]
            f = row_formulas(r)
            first = (po, inb) != prev
            prev = (po, inb)
            rows.append(["정기" if first else "", po if first else "", inb if first else "", brand, center, round(g["cost"]),
                         f[0], f[1], round(g["sales"]), f[3], f[4], f[5], f[6], f[7], f[8], f[9], f[10], f[11], f[12], 0,
                         AUTO_TAG_PROV if g["prov"] else AUTO_TAG_REAL])
        end = start + n - 1
        ws.update(f"A{start}:U{end}", rows, value_input_option="USER_ENTERED")
        # 서식은 수기 마지막 줄 서식을 복사
        sid = ws.id
        sh.batch_update({"requests": [{"copyPaste": {
            "source": {"sheetId": sid, "startRowIndex": manual_last - 1, "endRowIndex": manual_last, "startColumnIndex": 0, "endColumnIndex": 20},
            "destination": {"sheetId": sid, "startRowIndex": start - 1, "endRowIndex": end, "startColumnIndex": 0, "endColumnIndex": 20},
            "pasteType": "PASTE_FORMAT"}}, {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": start - 1, "endRowIndex": end, "startColumnIndex": 1, "endColumnIndex": 3},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "yyyy-mm-dd"}}},
            "fields": "userEnteredFormat.numberFormat"}}]})
        # 합계(10행)·브랜드별 마진(6행) 범위를 마지막 줄까지 넓힌다
        top = ws.get("A6:AG10", value_render_option="FORMULA")
        upd = []
        for ri, row in enumerate(top):
            for ci, v in enumerate(row):
                if isinstance(v, str) and v.startswith("=") and "11:" in v:
                    nv = re.sub(r"(\$?[A-Z]{1,2}\$?11:\$?[A-Z]{1,2}\$?)(\d+)", lambda m: m.group(1) + str(max(int(m.group(2)), end)), v)
                    if nv != v:
                        upd.append({"range": gspread.utils.rowcol_to_a1(6 + ri, ci + 1), "values": [[nv]]})
        if upd:
            ws.batch_update(upd, value_input_option="USER_ENTERED")
        report.append(f"{name} 탭 {start}~{end}행 ({n}줄{', 새 탭' if created else ''})")
    return report


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
    write_item_tab(sh, items)
    rep = write_month_tabs(sh, items)
    print(f"자동_품목 {len(items)}줄 갱신 · " + " / ".join(rep))
    if missing:
        print("LAW에 매입가 없음:", "; ".join(f"{c} {n}" for c, n in missing))
        open(os.environ.get("GITHUB_STEP_SUMMARY", os.devnull), "a").write("LAW에 매입가 없는 품목: " + ", ".join(n or c for c, n in missing) + "\n")


if __name__ == "__main__":
    main()
