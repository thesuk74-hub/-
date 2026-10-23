"""수량파일(설계내역서·수량집계표 엑셀) 해석 — 공종명·규격·단위·수량 행 추출

머리글 행(공종/품명/명칭 · 규격 · 단위 · 수량)을 시트별로 자동 탐지하여 수량 행을 읽는다.
  · 단위 없는 행(대공종 제목 '1. 토공' 등)은 현재 구분(제목)으로 기억하여 행별 분류에 사용
  · 소계·합계·계 행, 수량 0 행은 제외
  · 시트명·파일명·행 제목으로 공종 분류(토공·구조물공·가시설공·포장공·관로공·부대공·건축·기계·전기·기타)
"""
import os
import re

import openpyxl

CATEGORIES = ["토공", "가시설공", "구조물공", "관로공", "포장공", "부대공", "건축", "기계", "전기", "기타"]
CATEGORY_AUTO = "자동(행별 분류)"

NAME_KEYS = ("공종명", "공종", "품명", "명칭", "세부공종", "작업명", "공정명", "항목", "내역", "품목", "공사명")
SPEC_KEYS = ("규격", "사양", "규격및사양")
UNIT_KEYS = ("단위",)
QTY_KEYS = ("수량", "설계수량", "총수량", "물량", "합계수량", "계", "합계", "총계")
GROUP_KEYS = ("대공종", "공종구분", "구분", "분류")
SKIP_NAME = re.compile(r"^\s*(소\s*계|합\s*계|총\s*계|계|누\s*계|직접공사비|간접공사비|총\s*괄)\s*$|^\[?\s*(소계|합계)")

# 분류 규칙(위에서부터 우선) — 파일명·시트명·제목·공종명에 공통 적용
CAT_RULES = [
    ("시운전", r"시운전"),
    ("관로공", r"배수설비|오수받이|우수받이|빗물받이|연결관"),
    ("전기", r"전기|계장|계측제어|수변전|케이블|배선|조명|접지|CCTV|통신|EL\.|피뢰"),
    ("기계", r"기계|펌프|송풍기|교반기|탈수기|밸브류|설비|기기|산기|스크린|크레인|호이스트|덕트|공조|위생"),
    ("건축", r"건축|마감|조적|미장|도장|타일|창호|지붕|수장|내장|외장|벽돌|도배|석고"),
    ("가시설공", r"가시설|흙막이|토류|H-?PILE|천공|띠장|버팀보|시트\s*파일|SHEET|강널|복공|CIP|SCW|어스앵커|그라우팅|LW"),
    ("포장공", r"포장|아스콘|아스팔트|보조기층|택\s*코팅|프라임|절삭|보도블록|경계석|차선"),
    ("관로공", r"관로|관\s*부설|관\s*접합|부설|흄관|PE관|PVC관|덕타일|주철관|DCIP|VR관|하수관|오수관|우수관|배수관|맨홀|추진|연결관|받이|배수설비"),
    ("구조물공", r"구조물|콘크리트|레미콘|거푸집|유로폼|철근|동바리|비계|방수|타설|양생|지수판|기초|옹벽|박스|BOX|수조|PHC|말뚝"),
    ("토공", r"토공|터파기|굴착|되메우기|성토|절토|흙깎기|흙깍기|사토|토사|운반|다짐|모래|잡석|부설재|발파|리핑|암\s*깨기|줄파기"),
    ("부대공", r"부대|조경|울타리|안내|표지|정리|철거|가설|이설|보호공|지장물"),
]


def classify(*texts):
    t = " ".join(str(x or "") for x in texts)
    for cat, pat in CAT_RULES:
        if re.search(pat, t, re.I):
            return cat
    return None


def category_from_filename(fn):
    base = os.path.splitext(os.path.basename(fn))[0]
    c = classify(base)
    return c if c and c != "시운전" else CATEGORY_AUTO


def _norm(v):
    return re.sub(r"[\s　·_]", "", str(v or ""))


def _num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def _find_header(rows):
    """(머리글 행 번호, {name, spec, unit, qty, group}) — 2단 머리글(병합)도 위·아래 행을 합쳐 탐지."""
    for i, row in enumerate(rows[:40]):
        cells = [_norm(c) for c in row]
        nxt = [_norm(c) for c in rows[i + 1]] if i + 1 < len(rows) else []
        both = [a + b if a != b else a for a, b in zip(cells, nxt + [""] * (len(cells) - len(nxt)))] if nxt else cells
        for cand in (cells, both):
            col = {}
            for j, c in enumerate(cand):
                if not c:
                    continue
                if "unit" not in col and c in UNIT_KEYS:
                    col["unit"] = j
                elif "spec" not in col and any(c.startswith(k) for k in SPEC_KEYS):
                    col["spec"] = j
                elif "group" not in col and c in GROUP_KEYS:
                    col["group"] = j
                elif "name" not in col and any(c.startswith(k) or c.endswith(k) for k in NAME_KEYS):
                    col["name"] = j
            # 수량: '수량' 우선, 없으면 계/합계(금액 열 제외)
            for k in QTY_KEYS:
                js = [j for j, c in enumerate(cand) if c and (c == k or c.startswith(k)) and "금액" not in c and "단가" not in c]
                if js:
                    col["qty"] = js[0]
                    break
            if "unit" in col and "qty" in col and ("name" in col or "group" in col):
                if "name" not in col:
                    col["name"] = col.pop("group")
                hdr_end = i + (1 if cand is both and cand is not cells else 0)
                return hdr_end, col
    return None, None


def parse_sheet(ws):
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    h, col = _find_header(rows)
    if h is None:
        return None
    out, heading = [], ""
    for k, row in enumerate(rows[h + 1:], h + 2):
        def g(key):
            j = col.get(key)
            return row[j] if j is not None and j < len(row) else None
        name = str(g("name") or "").strip()
        grp = str(g("group") or "").strip()
        unit = str(g("unit") or "").strip()
        qty = _num(g("qty"))
        if not name and grp and not unit:
            heading = grp
            continue
        if not name:
            continue
        if SKIP_NAME.search(name):
            continue
        if not unit or qty is None:
            if not unit and len(name) <= 40:             # 제목행 (예: '1. 토공', '가. 구조물공')
                heading = re.sub(r"^[\s\d가-하ⅠⅡⅢⅣⅤ①-⑳()\.\-]+[\.\)]\s*", "", name) or name
            continue
        if qty <= 0:
            continue
        out.append({"row": k, "name": name, "spec": str(g("spec") or "").strip(), "unit": unit, "qty": qty,
                    "heading": grp or heading})
    return {"sheet": ws.title, "header_row": h + 1, "rows": out}


def parse_file(path):
    """반환: {"sheets": [{"sheet", "header_row", "rows": [...]}], "skipped": [시트명...]}"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheets, skipped = [], []
    for ws in wb.worksheets:
        try:
            r = parse_sheet(ws)
        except Exception:  # noqa: BLE001
            r = None
        if r and r["rows"]:
            sheets.append(r)
        else:
            skipped.append(ws.title)
    wb.close()
    return {"sheets": sheets, "skipped": skipped}


def row_category(file_cat, sheet, row):
    """행 분류: 파일 분류가 지정돼 있으면 그대로, 자동이면 시트명·제목·공종명 순으로 판단."""
    if file_cat and file_cat != CATEGORY_AUTO:
        return file_cat
    for t in (row.get("heading"), sheet, f"{row['name']} {row.get('spec', '')}"):
        c = classify(t)
        if c and c != "시운전":
            return c
    return "기타"
