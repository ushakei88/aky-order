#!/usr/bin/env python3
"""올리브영 발주 x 아케이 재고 매칭 → 암호화된 대시보드 데이터(data.enc, meta.json) 생성.

사용:
  DASH_PW=비번 python3 build.py --orders 발주.xlsx [--orders 발주2.xlsx] --stock 재고.xlsx [--master master.csv] [--out .]

- 발주 파일: 올리브원 납품확인서(R1) 엑셀 내려받기 형식
- 재고 파일: SCM '아케이 품목별 재고 & 출고 현황' 엑셀 (헤더 위치·열 이름을 자동으로 찾음)
- master.csv: 상품코드별 브랜드·소비기한일수·입수량·프리미엄(1/5 규칙)·발주메일
"""
import argparse, base64, csv, datetime as dt, hashlib, json, math, os, re, sys
from collections import defaultdict, OrderedDict

import openpyxl

KST = dt.timezone(dt.timedelta(hours=9))
ITER = 200000

# ---------------------------------------------------------------- 공통

def norm(s):
    return re.sub(r"[\s\(\)\[\]_/·.\-]", "", str(s or "")).lower()


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    if re.fullmatch(r"\d{8}", s):
        return dt.date(int(s[:4]), int(s[4:6]), int(s[6:]))
    m = re.match(r"(\d{2,4})[.\-/](\d{1,2})[.\-/](\d{1,2})", s)
    if m:
        y = int(m.group(1)); y = y + 2000 if y < 100 else y
        return dt.date(y, int(m.group(2)), int(m.group(3)))
    if isinstance(v, (int, float)) and 30000 < v < 80000:  # 엑셀 일련번호
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
    return None


def to_num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return v
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if "." in s else int(s)
    except ValueError:
        return None


def code_str(v):
    if v is None:
        return ""
    if isinstance(v, float):
        v = int(v)
    return re.sub(r"\D", "", str(v))


def find_header(ws, must, scan=15):
    """must: 헤더 행에 반드시 있어야 하는 키워드 목록. (행번호, {norm헤더: 열idx}) 반환"""
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=scan, values_only=True), start=1):
        cells = [norm(c) for c in row]
        if all(any(k in c for c in cells) for k in must):
            return r_idx, cells
    return None, None


def pick(cells, *keys, exclude=()):
    for k in keys:
        for i, c in enumerate(cells):
            if k in c and not any(x in c for x in exclude):
                return i
    return None


# 브랜드 통째로 고정 상태 (새 SKU가 떠도 자동 적용)
FIXED_BRANDS = {"GNC": "운영중지 (GNC 이관)"}


def brand_of(name):
    n = re.sub(r"^\[[^\]]*\]\s*", "", name or "")
    return n.split()[0] if n.split() else ""

# ---------------------------------------------------------------- 입력 읽기

def read_orders_json(p):
    """올리브원 화면(납품확인서 R1) 데이터셋을 그대로 뽑은 JSON"""
    d = json.load(open(p, encoding="utf-8"))
    out = []
    for r in d["rows"]:
        g = dict(zip(d["cols"], r))
        out.append({
            "slip": g.get("cntrPoSlipNo") or "", "date": to_date(g.get("stkinYmd")), "poDate": to_date(g.get("poYmd")),
            "center": g.get("stkinCntrCd") or "", "code": code_str(g.get("gdsCd")), "name": (g.get("gdsNm") or "").strip(),
            "box": to_num(g.get("inptUnitQty")) or 0, "ea": to_num(g.get("poQty")) or 0, "plan": to_num(g.get("supQty")),
            "mfg": to_date(g.get("mfgYmd")), "exp": to_date(g.get("dstrTlmtYmd")), "memo": (g.get("rmkCont") or "").strip(),
            "days": to_num(g.get("oymaster")) or None, "premium": g.get("premBrndYn") == "Y",
        })
    return out


def read_orders(paths):
    lines = []
    for p in paths:
        if p.lower().endswith(".json"):
            lines += read_orders_json(p)
            continue
        wb = openpyxl.load_workbook(p, data_only=True)
        ws = wb.active
        hr, cells = find_header(ws, ["상품코드", "발주수량"])
        if not hr:
            sys.exit(f"발주 파일 헤더를 찾지 못했어요: {p}")
        c = {
            "slip": pick(cells, "전표", exclude=("상세",)),
            "date": pick(cells, "입고예정일", "납품일", "입고일"),
            "center": pick(cells, "센터"),
            "code": pick(cells, "상품코드", "바코드"),
            "name": pick(cells, "상품명"),
            "box": pick(cells, "입력수량", "box"),
            "ea": pick(cells, "발주수량"),
            "plan": pick(cells, "납품예정수량"),
            "mfg": pick(cells, "제조일자"),
            "exp": pick(cells, "유통기한"),
            "memo": pick(cells, "제조일자와유통기한중", "메모", "비고"),
            "days": pick(cells, "소비기한일수", "유통일수"),
        }
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            code = code_str(row[c["code"]]) if c["code"] is not None else ""
            if not code:
                continue
            g = lambda k: row[c[k]] if c[k] is not None and c[k] < len(row) else None
            center = str(g("center") or "").strip()
            box = to_num(g("box")) or 0
            ea = to_num(g("ea")) or 0
            lines.append({
                "slip": str(g("slip") or ""),
                "date": to_date(g("date")),
                "center": center,
                "code": code,
                "name": str(g("name") or "").strip(),
                "box": box,
                "ea": ea,
                "plan": to_num(g("plan")),
                "mfg": to_date(g("mfg")),
                "exp": to_date(g("exp")),
                "memo": str(g("memo") or "").strip(),
                "days": to_num(g("days")),
            })
    return lines


def read_stock(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "올영" in wb.sheetnames:
        return read_stock_scm(wb)
    return read_stock_generic(path)


def read_stock_scm(wb):
    """SCM '아케이 품목별 재고 & 출고 현황' 형식: 올영(올리브영 전용 재고) + 올영EXP + EXP(일반재고 유통기한, 상품명 매칭)"""
    ws = wb["올영"]
    rows = list(ws.iter_rows(min_row=1, max_row=12, values_only=True))
    hr = next(i for i, r in enumerate(rows) if any(norm(c) == "상품코드" for c in r))
    cells = [norm(c) for c in rows[hr]]
    cc, nc = cells.index("상품코드"), pick(cells, "상품명")
    qc, uc = pick(cells, "현재고"), pick(cells, "카톤입수", "입수")
    out = OrderedDict()
    byname = {}
    for r in ws.iter_rows(min_row=hr + 2, values_only=True):
        if cc >= len(r):
            continue
        code = code_str(r[cc])
        if len(code) < 8:
            continue
        it = out.setdefault(code, {"qty": 0, "name": str(r[nc] or "").strip(), "lots": [], "unit": None, "expref": False})
        it["qty"] += to_num(r[qc]) or 0
        u = to_num(r[uc]) if uc is not None else None
        if u:
            it["unit"] = int(u)
        byname[re.sub(r"\s", "", it["name"])] = code
    nlots = 0
    if "올영EXP" in wb.sheetnames:
        for r in wb["올영EXP"].iter_rows(min_row=1, values_only=True):
            code = code_str(r[0]) if r and r[0] else ""
            exp = to_date(r[3]) if len(r) > 3 else None
            if code in out and exp:
                out[code]["lots"].append({"exp": exp, "mfg": None, "qty": to_num(r[4]) or 0})
                nlots += 1
    nref = 0
    if "EXP" in wb.sheetnames:
        for r in wb["EXP"].iter_rows(min_row=1, values_only=True):
            if not r or len(r) < 5:
                continue
            code = byname.get(re.sub(r"\s", "", str(r[2] or "")))
            exp = to_date(r[3])
            if code and exp and not any(l["qty"] for l in out[code]["lots"]):
                out[code]["lots"].append({"exp": exp, "mfg": None, "qty": None})
                out[code]["expref"] = True
                nref += 1
    # 향동(자체물류) 재고: '요약' 탭의 재고 > 향동자체물류 칸. 청담소녀처럼 향동 재고를 이관해 쓰는 브랜드용
    hyang = {}
    if "요약" in wb.sheetnames:
        rs = list(wb["요약"].iter_rows(max_row=2000, values_only=True))
        hi = next((i for i, r in enumerate(rs[:20]) if r and "품목명" in [str(c).strip() if c else "" for c in r]), None)
        if hi is not None:
            h = [str(c).strip() if c else "" for c in rs[hi]]
            ci, ni = h.index("품목코드"), h.index("품목명")
            qi = h.index("향동자체물류") if "향동자체물류" in h else None
            if qi is not None:
                for r in rs[hi + 1:]:
                    if r and len(r) > qi and r[ni]:
                        hyang[str(r[ni]).strip()] = {"qty": to_num(r[qi]) or 0, "icode": str(r[ci] or "")}
    meta = {"sheet": "올영", "qtyCol": "현재고", "rows": len(out), "oyExpLots": nlots, "refExpLots": nref, "hyang": len(hyang)}
    out["__hyang__"] = hyang
    return out, meta


def read_stock_generic(path):
    """반환: {code: {"qty": 총재고, "name": 상품명, "lots": [{"exp": date|None, "mfg": date|None, "qty": n}]}}"""
    wb = openpyxl.load_workbook(path, data_only=True)
    best = None
    for ws in wb.worksheets:
        for keys in (["상품코드"], ["바코드"], ["상품바코드"], ["sku"]):
            hr, cells = find_header(ws, keys, scan=20)
            if hr:
                q = pick(cells, "가용재고", "현재고", "정상재고", "재고수량", "재고", "잔량", exclude=("출고", "입고", "일수", "금액"))
                if q is not None:
                    best = (ws, hr, cells, q)
                    break
        if best:
            break
    if not best:
        sys.exit("재고 파일에서 '상품코드/바코드'와 '재고' 열을 찾지 못했어요.")
    ws, hr, cells, qcol = best
    ccol = pick(cells, "상품코드", "상품바코드", "바코드", "sku")
    ncol = pick(cells, "상품명", "품목명", "품명")
    ecol = pick(cells, "유통기한", "소비기한", exclude=("일수",))
    mcol = pick(cells, "제조일자", "제조일")
    out = OrderedDict()
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        if ccol >= len(row):
            continue
        code = code_str(row[ccol])
        if len(code) < 8:
            continue
        qty = to_num(row[qcol]) or 0
        it = out.setdefault(code, {"qty": 0, "name": str(row[ncol] or "").strip() if ncol is not None else "", "lots": []})
        it["qty"] += qty
        exp = to_date(row[ecol]) if ecol is not None else None
        mfg = to_date(row[mcol]) if mcol is not None else None
        if exp or mfg:
            it["lots"].append({"exp": exp, "mfg": mfg, "qty": qty})
    meta = {"sheet": ws.title, "qtyCol": cells[qcol], "hasExp": ecol is not None, "hasMfg": mcol is not None, "rows": len(out)}
    return out, meta


def read_master(path):
    m = {}
    if path and os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                code = code_str(r.get("상품코드"))
                if code:
                    m[code] = r
    return m

# ---------------------------------------------------------------- 계산

def center_type(center):
    c = center.upper()
    if "LA04" in c or "온라인" in center:
        return "온라인"
    return "오프라인"


CENTER_NAME = {"LA02": "양지센터", "LA04": "양지On센터"}


def analyze(lines, stock, master, prev):
    by = OrderedDict()
    for ln in lines:
        by.setdefault(ln["code"], []).append(ln)

    delivery = None
    ds = [ln["date"] for ln in lines if ln["date"]]
    if ds:
        delivery = max(set(ds), key=ds.count)

    items = []
    for code, ls in by.items():
        ms = master.get(code, {})
        st = stock.get(code)
        name = ls[0]["name"]
        total = sum(l["ea"] for l in ls)
        boxes = sum(l["box"] for l in ls)
        unit = None
        for l in ls:
            if l["box"]:
                unit = round(l["ea"] / l["box"])
                break
        unit = int(to_num(ms.get("입수량")) or (st or {}).get("unit") or unit or 1)
        days = int(next((l["days"] for l in ls if l["days"]), 0) or to_num(ms.get("소비기한일수")) or 0)
        premium = str(ms.get("프리미엄", "")).strip() in ("1", "Y", "y", "O", "o", "예") or any(l.get("premium") for l in ls)
        frac = 5 if premium else 3
        qty = int(st["qty"]) if st else 0

        issues = []
        if st and st.get("hyMissing"):
            issues.append({"lv": "warn", "t": "향동 재고표(요약 탭)에서 이 품목을 못 찾았어요 — master.csv '향동품목코드' 칸에 품목코드를 적어 주세요"})
        lots = []
        eligible = None
        if delivery and days:
            src_lots = (st["lots"] if st and st["lots"] else [])
            if not src_lots:
                # 재고 파일에 유통기한이 없으면 R1에 입력된 유통기한으로 검사
                src_lots = [{"exp": l["exp"], "mfg": l["mfg"], "qty": None} for l in ls if l["exp"] or l["mfg"]]
            eligible = 0 if src_lots and any(x["qty"] is not None for x in src_lots) else None
            seen = set()
            for lot in src_lots:
                exp = lot["exp"] or (lot["mfg"] + dt.timedelta(days=days - 1) if lot["mfg"] else None)
                if not exp or (exp, lot["qty"]) in seen:
                    continue
                seen.add((exp, lot["qty"]))
                remain = (exp - delivery).days
                mfg_calc = exp - dt.timedelta(days=days - 1)
                future = mfg_calc > delivery
                ok = remain > days / frac and not future
                ban = exp - dt.timedelta(days=days // 2)
                lots.append({"exp": exp.isoformat(), "qty": lot["qty"], "remain": remain, "pct": round(remain / days * 100),
                             "ok": ok, "future": future, "half": remain <= days / 2})
                if lot["qty"] is not None and ok:
                    eligible += lot["qty"]
                if future:
                    issues.append({"lv": "bad", "t": f"유통기한 {exp:%y.%m.%d} → 계산 제조일 {mfg_calc:%y.%m.%d}가 납품일보다 미래라 납품 불가 (임의 입력·수기수정 시 회송됨)"})
                elif not ok:
                    issues.append({"lv": "bad", "t": f"유통기한 {exp:%y.%m.%d} 잔여 {remain}일({round(remain/days*100)}%) — 1/{frac} 이하라 납품 불가"})
                elif remain <= days / 2:
                    issues.append({"lv": "note", "t": f"유통기한 {exp:%y.%m.%d} 잔여 {round(remain/days*100)}% — 사내 입고금지일({ban:%y.%m.%d}) 지남"})
            # 선입선출로 발주수량을 채우는 데 필요한 유통기한 종류 수
            ok_lots = sorted([l for l in lots if l["ok"]], key=lambda l: l["exp"])
            need, acc, kinds = total, 0, []
            for l in ok_lots:
                if acc >= need:
                    break
                if l["qty"] is None or l["qty"] > 0:
                    kinds.append(l["exp"]); acc += l["qty"] or need
            ok_kinds = len(set(kinds))
            if ok_kinds == 2:
                issues.append({"lv": "warn", "t": "유통기한 2가지 — G코드 예외등록 필요 (대표님 승인)"})
            elif ok_kinds >= 3:
                issues.append({"lv": "bad", "t": f"유통기한 {ok_kinds}가지 — 납품 불가"})
        elif not days:
            issues.append({"lv": "note", "t": "소비기한일수 정보 없음 — 기준표에 추가 필요"})

        if st and st.get("expref") and lots:
            issues.append({"lv": "note", "t": "유통기한은 SCM 일반재고(EXP 시트) 기준 참고값이에요 — 올리브영 전용 재고 유통기한은 3PL 회신으로 확인"})
        usable = qty if eligible is None else min(qty, eligible)
        after = usable - total
        short = max(0, -after)
        cartons = math.ceil(short / unit) if short else 0
        memo = " / ".join(sorted({l["memo"] for l in ls if l["memo"]}))
        stop = any(k in memo for k in ("미운영", "발주중지", "미납"))

        fixed = str(ms.get("고정상태", "")).strip()
        if not fixed and (ms.get("브랜드") or brand_of(name)) in FIXED_BRANDS:
            fixed = FIXED_BRANDS[ms.get("브랜드") or brand_of(name)]
        st_need, cartons, why = need_rule(after, unit)
        if fixed:
            status, cartons, why = "end", 0, fixed
        elif stop:
            status, cartons = "stop", 0
        elif not st and st_need in ("short",):
            status = "nostock"
        elif st_need in ("short", "pre"):
            status = st_need
        elif any(i["lv"] == "bad" for i in issues):
            status = "bad"
        else:
            status = "ok"

        p = (prev or {}).get(code)
        change = None
        if prev is not None:
            if p is None:
                change = "new"
            elif p != total:
                change = "qty"

        items.append({
            "code": code, "name": name, "brand": ms.get("브랜드") or brand_of(name),
            "mail": ms.get("발주메일", ""),
            "lines": [{"center": CENTER_NAME.get(l["center"], l["center"]), "type": center_type(l["center"]),
                       "box": l["box"], "ea": l["ea"], "slip": l["slip"]} for l in ls],
            "online": sum(l["ea"] for l in ls if center_type(l["center"]) == "온라인"),
            "offline": sum(l["ea"] for l in ls if center_type(l["center"]) == "오프라인"),
            "total": total, "boxes": boxes, "unit": unit, "stock": qty, "src": (st or {}).get("src", "올영"), "usable": usable,
            "after": after, "short": short, "cartons": cartons, "orderEA": cartons * unit,
            "days": days, "frac": frac, "lots": lots, "issues": issues, "memo": memo, "stop": stop,
            "status": status, "change": change, "prevTotal": p, "why": why,
        })

    apply_bundles(items, stock, master)

    order = {"short": 0, "nostock": 1, "bad": 2, "pre": 3, "ok": 4, "stop": 5, "end": 6}
    items.sort(key=lambda x: (order[x["status"]], -x["short"], x["brand"], x["name"]))

    removed = []
    if prev:
        for code, ea in prev.items():
            if code not in by:
                removed.append({"code": code, "prevTotal": ea, "name": (prev_names or {}).get(code, "")})
    return delivery, items, removed


def need_rule(after, unit):
    """납품 후 남는 수량 기준 발주 판단: (상태, 권장카톤, 이유)"""
    if after < 0:
        return "short", math.ceil(-after / unit), f"부족 {-after:,}개"
    if after == 0:
        return "short", 1, "납품하면 재고 0 — 1카톤 발주"
    if after <= unit:
        return "pre", 1, f"납품 후 {after:,}개 남음(1카톤 이하) — 미리 1카톤 발주"
    return "ok", 0, ""


BUNDLE_BRANDS = ("청담소녀",)


def base_key(name):
    n = re.sub(r"\*\s*\d+", "", name or "")
    n = re.sub(r"\((온|역|오프)\)", "", n)
    return re.sub(r"\s", "", n)


HYANGDONG_BRANDS = ("청담소녀",)  # 올영 재고가 아니라 향동 재고를 이관해서 쓰는 브랜드


def apply_hyangdong(stock, master, lines):
    hy = stock.pop("__hyang__", {}) or {}
    if not hy:
        return
    by_key = {base_key(n): v for n, v in hy.items()}
    by_icode = {v["icode"]: v for v in hy.values() if v.get("icode")}
    cands = {}
    for c, r in master.items():
        cands[c] = r.get("상품명") or ""
    for l in lines:
        cands.setdefault(l["code"], l["name"])
    for c, st in list(stock.items()):
        cands.setdefault(c, st.get("name") or "")
    for c, nm in cands.items():
        if not any(b in nm for b in HYANGDONG_BRANDS) or "*" in nm:
            continue  # 번들(*3 등)은 단품 재고로 계산
        ic = (master.get(c) or {}).get("향동품목코드")
        v = by_icode.get(ic) if ic else by_key.get(base_key(nm))
        stock[c] = {"qty": (v or {}).get("qty", 0), "name": nm, "lots": (stock.get(c) or {}).get("lots", []), "unit": (stock.get(c) or {}).get("unit"),
                    "expref": False, "src": "향동", "hyMissing": v is None}


def links_of(master, items, stock):
    """묶음 관계: 묶음/여러개입 상품코드 → (단품코드, 배수)
    1) master.csv '단품코드'·'배수' 칸  2) 청담소녀 '…*3' 같은 이름 규칙"""
    links = {}
    for code, r in master.items():
        b, k = code_str(r.get("단품코드")), to_num(r.get("배수"))
        if b and k:
            links[code] = (b, int(k))
    names = {}
    for code, st in stock.items():
        names.setdefault(base_key(st.get("name")), code)
    for code, r in master.items():
        if "*" not in (r.get("상품명") or ""):
            names.setdefault(base_key(r.get("상품명")), code)
    for it in items:
        if "*" not in it["name"]:
            names.setdefault(base_key(it["name"]), it["code"])
    for it in items:
        if it["code"] in links or not any(b in it["name"] for b in BUNDLE_BRANDS):
            continue
        m = re.search(r"\*\s*(\d+)", it["name"])
        if m and names.get(base_key(it["name"])):
            links[it["code"]] = (names[base_key(it["name"])], int(m.group(1)))
    return links


def apply_bundles(items, stock, master):
    """묶음은 단품 기준으로 합쳐서 계산.
    - 청담소녀 *3·*4: 단품을 묶어서 번들로 출고 → 번들 발주 × N 을 단품 재고에서 뺀다
    - 요헤미티 7포 등: 7포를 뜯어 1포로도 출고 → 7포 재고 × 7 을 단품 재고로 친다
    둘 다 '단품 수량'으로 바꾸면 같은 계산: 가능 = Σ재고×배수, 필요 = Σ발주×배수"""
    links = links_of(master, items, stock)
    rev = {}
    for c, (b, k) in links.items():
        rev.setdefault(b, []).append((c, k))
    groups = OrderedDict()
    for it in items:
        b, k = links.get(it["code"], (it["code"], 1))
        it["mult"], it["base"] = k, b
        groups.setdefault(b, []).append(it)
    for base, mem in groups.items():
        linked = rev.get(base, [])
        if len(mem) == 1 and mem[0]["mult"] == 1 and not linked:
            continue
        members = {base: 1, **{c: k for c, k in linked}}
        parts, avail = [], 0
        for c, k in members.items():
            q = int(stock.get(c, {}).get("qty", 0))
            avail += q * k
            nm = (stock.get(c, {}).get("name") or (master.get(c) or {}).get("상품명") or next((m["name"] for m in mem if m["code"] == c), c))
            if q:
                parts.append(f"{short_name(nm)} {q:,}" + (f"×{k}" if k > 1 else ""))
        demand = sum(m["total"] * m["mult"] for m in mem)
        after = avail - demand
        short = max(0, -after)
        single = next((m for m in mem if m["code"] == base), None)
        bname = (single or {}).get("name") or stock.get(base, {}).get("name") or (master.get(base) or {}).get("상품명") or base
        unit = int(to_num((master.get(base) or {}).get("입수량")) or (single or {}).get("unit") or stock.get(base, {}).get("unit") or 1)
        detail = " + ".join(f"{short_name(m['name'])} {m['total']}" + (f"×{m['mult']}" if m["mult"] > 1 else "") for m in mem)
        src = stock.get(base, {}).get("src")
        rep = single or mem[0]
        for m in mem:
            m["after"] = after if m is rep else 0
            m["short"] = m["cartons"] = m["orderEA"] = 0
            m["stock"] = m["usable"] = avail      # 화면에는 단품 기준으로 쓸 수 있는 수량
            m["src"] = src or "올영"
            m["pooled"] = True
            m["poolNote"] = f"단품 기준 계산: 필요 {demand:,}개 ({detail}) · 쓸 수 있는 단품 {avail:,}개 ({' + '.join(parts) or '재고 없음'}{', 향동 재고' if src == '향동' else ''})"
            m["issues"].insert(0, {"lv": "note", "t": m["poolNote"] + ("" if m is rep else f" → 발주는 '{bname}'에 합산")})
        st_need, ctn, why = need_rule(after, unit)
        for m in mem:
            if m["status"] not in ("stop", "end"):
                m["status"] = st_need if st_need != "ok" else ("bad" if any(i["lv"] == "bad" for i in m["issues"]) else "ok")
                m["why"] = why
        rep["short"] = short
        rep["cartons"] = ctn if rep["status"] not in ("stop", "end") else 0
        rep["orderEA"] = rep["cartons"] * unit
        rep["unit"] = unit
        rep["orderName"], rep["orderCode"] = bname, base


def short_name(n):
    n = re.sub(r"\((온|역)\)", "", n or "").strip()
    return n.split(" ", 1)[1] if " " in n else n


prev_names = {}

# ---------------------------------------------------------------- 암호화

def kdf(pw, salt):
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(pw.encode())


def encrypt(pw, salt, data: bytes):
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    iv = os.urandom(12)
    return iv + AESGCM(kdf(pw, salt)).encrypt(iv, data, None)


def decrypt_prev(pw, outdir, name="data.enc"):
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        meta = json.load(open(os.path.join(outdir, "meta.json")))
        buf = open(os.path.join(outdir, name), "rb").read()
        key = kdf(pw, base64.b64decode(meta["salt"]))
        return json.loads(AESGCM(key).decrypt(buf[:12], buf[12:], None))
    except Exception:
        return None

# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--orders", action="append", required=True)
    ap.add_argument("--stock")
    ap.add_argument("--master", default=os.path.join(os.path.dirname(__file__), "master.csv"))
    ap.add_argument("--out", default=os.path.dirname(os.path.abspath(__file__)))
    ap.add_argument("--stock-label", default="")
    a = ap.parse_args()
    pw = os.environ.get("DASH_PW")
    if not pw:
        sys.exit("DASH_PW 환경변수(대시보드 비밀번호)가 필요해요.")

    lines = read_orders(a.orders)
    stock, smeta = read_stock(a.stock) if a.stock else ({}, {"rows": 0})
    master = read_master(a.master)
    apply_hyangdong(stock, master, lines)

    ds = [l["date"] for l in lines if l["date"]]
    delivery_guess = max(set(ds), key=ds.count) if ds else None
    dkey = delivery_guess.isoformat() if delivery_guess else None
    today = dt.datetime.now(KST).date()
    arch_name = f"archive/{dkey}.enc" if dkey else None
    if delivery_guess and delivery_guess < today and arch_name and os.path.exists(os.path.join(a.out, arch_name)):
        print(f"납품일 {dkey}은 지난 날짜라 기록을 고정해 둬요 (덮어쓰지 않음).")
        return
    old = decrypt_prev(pw, a.out, arch_name) if arch_name and os.path.exists(os.path.join(a.out, arch_name)) else decrypt_prev(pw, a.out)
    prev = None
    if old and old.get("delivery") == dkey:
        prev = {i["code"]: i["total"] for i in old["items"]}
        prev_names.update({i["code"]: i["name"] for i in old["items"]})
    delivery, items, removed = analyze(lines, stock, master, prev)

    now = dt.datetime.now(KST)
    hist = (old or {}).get("history", []) if prev is not None else []
    sig = hashlib.sha1(json.dumps([(i["code"], i["total"]) for i in items]).encode()).hexdigest()[:10]
    changed = [i for i in items if i["change"]]
    if not hist or hist[-1].get("sig") != sig:
        hist.append({"at": now.strftime("%Y-%m-%d %H:%M"), "sig": sig, "sku": len(items), "ea": sum(i["total"] for i in items),
                     "new": sum(1 for i in changed if i["change"] == "new"), "qty": sum(1 for i in changed if i["change"] == "qty"),
                     "removed": len(removed)})

    # 발주일(처음 계산한 날) 재고를 기억해 두고, 그 뒤 들어온 수량을 +로 보여준다
    base_label = None
    if prev is not None and old:
        ob = {i["code"]: i for i in old.get("items", [])}
        base_label = old.get("baseLabel") or old.get("stockLabel")
        for i in items:
            o = ob.get(i["code"])
            same_src = o and o.get("src", "올영") == i.get("src", "올영")  # 재고 기준(올영/향동)이 바뀌면 기준점도 새로
            i["baseStock"] = o.get("baseStock", o.get("stock")) if same_src else i["stock"]
    else:
        for i in items:
            i["baseStock"] = i["stock"]
    for i in items:
        i["inbound"] = i["stock"] - (i["baseStock"] or 0)

    data = {
        "baseLabel": base_label or (a.stock_label or (os.path.basename(a.stock) if a.stock else "")),
        "built": now.strftime("%Y-%m-%d %H:%M"),
        "delivery": delivery.isoformat() if delivery else None,
        "stockLabel": a.stock_label or (os.path.basename(a.stock) if a.stock else "재고 파일 없음"),
        "stockMeta": smeta,
        "items": items, "removed": removed, "history": hist[-30:],
        "brandMail": {r.get("브랜드"): r.get("발주메일") for r in master.values() if r.get("브랜드") and r.get("발주메일")},
    }
    try:  # 회송 목록 (납품확인서 검수 확정 + CJ 납품완료보고서 사유)
        from returns_build import build_returns
        data["returns"] = build_returns(pw)
    except Exception as e:
        print("회송 목록 건너뜀:", e)
        data["returns"] = (old or {}).get("returns", [])
    raw = json.dumps(data, ensure_ascii=False, default=str).encode()

    meta_path = os.path.join(a.out, "meta.json")
    salt = None
    if os.path.exists(meta_path):
        try:
            m = json.load(open(meta_path))
            if decrypt_prev(pw, a.out) is not None:
                salt = base64.b64decode(m["salt"])
        except Exception:
            pass
    salt = salt or os.urandom(16)
    enc = encrypt(pw, salt, raw)
    if data["delivery"]:
        os.makedirs(os.path.join(a.out, "archive"), exist_ok=True)
        open(os.path.join(a.out, f"archive/{data['delivery']}.enc"), "wb").write(enc)
    # 대시보드 첫 화면 = 오늘 이후 가장 가까운 납품일 (없으면 방금 만든 것)
    dates = sorted(f[:-4] for f in os.listdir(os.path.join(a.out, "archive")) if f.endswith(".enc")) if os.path.isdir(os.path.join(a.out, "archive")) else []
    upcoming = [d for d in dates if d >= today.isoformat()]
    current = data["delivery"] if (not upcoming or data["delivery"] == upcoming[0]) else upcoming[0]
    if current == data["delivery"]:
        open(os.path.join(a.out, "data.enc"), "wb").write(enc)
    json.dump({"salt": base64.b64encode(salt).decode(), "iter": ITER, "built": data["built"], "current": current, "dates": dates},
              open(meta_path, "w"), ensure_ascii=False)

    s = defaultdict(int)
    for i in items:
        s[i["status"]] += 1
    print(f"납품일 {data['delivery']} · {len(items)} SKU · 발주필요 {s['short']} · 재고없음 {s['nostock']} · 미리발주 {s['pre']} · 유통기한불가 {s['bad']} · 충분 {s['ok']} · 운영중지 {s['end']} · 총 {sum(i['cartons'] for i in items)}카톤")
    print(f"변동: 신규 {sum(1 for i in changed if i['change']=='new')} · 수량변경 {sum(1 for i in changed if i['change']=='qty')} · 삭제 {len(removed)}")
    print("재고 파일:", smeta)


if __name__ == "__main__":
    main()
