"""수량첨부 검토 API — ① 공정 선택·수량파일 첨부, ② 2.1 공사개요 ~ 2.6 CP 검토 (웹 화면 index.html)

작업 데이터: web_data/<작업ID>/입력/수량/<공정명>/<파일>, meta.json["wiz"] = wizard.state
검토 실행·보고서 작성·출력·결과·파일은 app.py 공통 API(/run, /report, /export, /result, /file) 사용.
"""
import datetime as dt
import os
import shutil
from typing import List

from fastapi import APIRouter, File, Form, HTTPException, UploadFile

from .. import refdata as R
from .. import wizard as W
from ..nonwork import NonworkSetup
from ..project import FORMULA_VAR_KEYS, OVERVIEW
from ..qtyparse import CATEGORIES, CATEGORY_AUTO, category_from_filename

router = APIRouter()
CTX = {}   # app.py 에서 주입: jdir, meta, save_meta, new_job_dir, rt


def _m(jid):
    m = CTX["meta"](jid)
    if m.get("type") != "수량첨부":
        raise HTTPException(400, "수량첨부 방식 작업이 아닙니다.")
    m.setdefault("wiz", W.default_state())
    return m


def base_dir(jid):
    return os.path.join(CTX["jdir"](jid), "입력")


def input_path(jid):
    return os.path.join(base_dir(jid), "검토입력.xlsx")


def build(jid):
    """state → 검토입력.xlsx (검토 실행 직전·조회 시)."""
    m = _m(jid)
    if not m["wiz"]["works"] or not any(w["files"] for w in m["wiz"]["works"]):
        raise HTTPException(400, "수량파일이 없습니다. [① 파일 추가]에서 공정 선택 후 수량파일을 첨부하십시오.")
    return W.build_input(m["wiz"], base_dir(jid), input_path(jid))


def _save(jid, m):
    CTX["save_meta"](jid, m)
    CTX["rt"](jid)["res"] = None


# ── 기준정보 ──────────────────────────────────────────────────────────────────
@router.get("/api/wiz/info")
def info():
    opts = {k: o for k, _, _, o in OVERVIEW if o}
    regions = {}
    for k, v in R.WEATHER["stations"].items():
        regions.setdefault(v["region"], []).append(k)
    f5 = {k: {"식": v[4], "변수": v[1], "범위": v[3], "vars": R.FORMULA_VARS.get(k, [])} for k, v in R.FORMULAS.items()}
    keco = {k: {"식": v[1], "변수": v[2], "출처": v[3], "vars": ["V"]} for k, v in R.KECO_FORMULAS.items()}
    exp = {k: {"표": " / ".join(f"{a}억 이하 {b}개월" for a, b in v), "출처": R.KECO_EXP_SOURCE} for k, v in R.KECO_EXPERIENCE.items()}
    return {
        "tree": {g: list(sub) for g, sub in R.WORK_TREE.items()}, "work_info": R.WORK_INFO,
        "categories": CATEGORIES, "category_auto": CATEGORY_AUTO,
        "corrections": {k: {"일수": v[0], "대상": v[1], "설명": v[2]} for k, v in R.CORRECTION_ITEMS.items()},
        "correction_source": R.CORRECTION_SOURCE, "indirect_source": R.INDIRECT_SOURCE,
        "indirect_rules": [{"대상": p, "간접작업": w, "일수": d} for p, w, d in R.INDIRECT_RULES],
        "regions": {k: sorted(v) for k, v in sorted(regions.items())},
        "conditions": {c: {"name": R.CONDITIONS[c], "short": R.COND_SHORT[c], "group": R.COND_GROUP[c]} for c in range(1, 16)},
        "profiles": {k: {str(c): w for c, w in v.items()} for k, v in R.DEFAULT_PROFILES.items()}, "profile_desc": R.PROFILE_DESC,
        "dust": R.DUST_EXAMPLES, "formulas": f5, "keco": keco, "keco_exp": exp,
        "var_keys": {v: k for k, v in FORMULA_VAR_KEYS.items()},
        "options": {k: opts.get(k) for k in ("현장여건", "공사규모", "지질조건", "기상기후", "준비기간 공사유형", "동절기 공사중지",
                                             "목표기간 기준", "적정성 판정기준", "시설유형(부록5)", "작업일수 산정단위(일)")},
        "prep_periods": R.PREP_PERIODS, "guideline": R.GUIDELINE,
        "defaults": {**{k: (str(d) if d is not None else None) for k, d, desc, _ in OVERVIEW if desc is not None},
                     "목표기간 기준": "준공예정일"},   # 수량첨부 검토: 공사개요 준공예정일을 목표기간으로
        "db": [[it["no"], it["field"], it["group"], it["name"], it["condition"], it["rate"], it["unit"]] for it in R.PRODUCTIVITY["items"]],
    }


# ── 작업 생성·조회 ────────────────────────────────────────────────────────────
@router.post("/api/wiz")
def create(body: dict):
    jid, d = CTX["new_job_dir"]()
    os.makedirs(os.path.join(d, "입력", "수량"), exist_ok=True)
    st = W.default_state()
    works = [w for w in body.get("works", []) if w in R.WORK_INFO]
    st["works"] = [{"name": w, "files": []} for w in works]
    name = (body.get("name") or "").strip() or "수량첨부 검토"
    st["ov"]["사업명"] = name
    CTX["save_meta"](jid, {"name": name, "type": "수량첨부", "kind": "공정선택·수량첨부", "wiz": st,
                           "inputs": ["검토입력.xlsx"], "current": "검토입력.xlsx",
                           "created": dt.datetime.now().strftime("%Y-%m-%d %H:%M"), "history": []})
    return {"id": jid}


def _files_info(jid, st):
    out = []
    for w in st["works"]:
        fl = []
        for i, f in enumerate(w["files"]):
            try:
                pr = W.parsed(base_dir(jid), f)
                sheets = [{"sheet": s["sheet"], "rows": len(s["rows"]), "header_row": s["header_row"],
                           "on": s["sheet"] not in (f.get("sheets_off") or [])} for s in pr["sheets"]]
                err = "" if sheets else "수량 머리글(공종·규격·단위·수량)을 찾지 못했습니다"
            except Exception as e:  # noqa: BLE001
                sheets, err = [], f"읽기 오류: {e}"
            fl.append({"idx": i, "orig": f.get("orig"), "category": f.get("category"), "sheets": sheets, "error": err,
                       "rows": sum(s["rows"] for s in sheets if s["on"])})
        out.append({"name": w["name"], "group": R.WORK_INFO.get(w["name"], {}).get("group"),
                    "code": R.WORK_INFO.get(w["name"], {}).get("code"), "files": fl})
    return out


def _suggest(jid, st):
    try:
        return W.suggest_formulas(st, W.collect(st, base_dir(jid)))
    except Exception:  # noqa: BLE001
        return W.suggest_formulas(st)


@router.get("/api/jobs/{jid}/wiz")
def get_state(jid: str):
    m = _m(jid)
    st = m["wiz"]
    ov = st["ov"]
    station, why = W.auto_station(ov.get("사업위치"))
    days = None
    try:
        if ov.get("착공예정일") and ov.get("준공예정일"):
            days = (dt.date.fromisoformat(ov["준공예정일"][:10]) - dt.date.fromisoformat(ov["착공예정일"][:10])).days + 1
    except ValueError:
        pass
    return {"id": jid, "name": m["name"], "state": st, "works": _files_info(jid, st),
            "auto": {"station": station, "station_reason": why, "dust": W.auto_dust(ov.get("기상관측지점") or station),
                     "formulas": _suggest(jid, st)},
            "target_days": days, "target_months": round(days / R.DAYS_PER_MONTH, 1) if days else None,
            "summary": m.get("summary")}


@router.post("/api/jobs/{jid}/wiz/works")
def set_works(jid: str, body: dict):
    m = _m(jid)
    want = [w for w in body.get("works", []) if w in R.WORK_INFO]
    cur = {w["name"]: w for w in m["wiz"]["works"]}
    for w in cur:
        if w not in want and cur[w]["files"] and not body.get("force"):
            raise HTTPException(400, f"'{w}'에 첨부된 수량파일이 있습니다. 파일을 삭제한 뒤 선택 해제하십시오.")
    m["wiz"]["works"] = [cur.get(w) or {"name": w, "files": []} for w in want]
    _save(jid, m)
    return {"ok": True}


@router.post("/api/jobs/{jid}/wiz/files")
async def add_files(jid: str, work: str = Form(...), files: List[UploadFile] = File(...)):
    m = _m(jid)
    w = next((x for x in m["wiz"]["works"] if x["name"] == work), None)
    if not w:
        raise HTTPException(400, f"선택되지 않은 공정입니다: {work}")
    dest = os.path.join(base_dir(jid), "수량", work.replace("/", "_"))
    os.makedirs(dest, exist_ok=True)
    added, errs = [], []
    for f in files:
        fn = os.path.basename(f.filename or "수량.xlsx")
        if not fn.lower().endswith((".xlsx", ".xlsm")):
            errs.append(f"'{fn}': 엑셀(.xlsx) 파일만 첨부할 수 있습니다.")
            continue
        p = os.path.join(dest, fn)
        base, ext = os.path.splitext(p)
        k = 2
        while os.path.exists(p):
            p = f"{base}_{k}{ext}"
            k += 1
        with open(p, "wb") as fh:
            fh.write(await f.read())
        ent = {"file": os.path.relpath(p, base_dir(jid)).replace(os.sep, "/"), "orig": os.path.basename(p),
               "category": category_from_filename(fn), "sheets_off": []}
        try:
            pr = W.parsed(base_dir(jid), ent)
            if not pr["sheets"]:
                errs.append(f"'{fn}': 수량 머리글(공종·규격·단위·수량)을 찾지 못했습니다 — 첨부는 유지, 파일 확인 필요")
        except Exception as e:  # noqa: BLE001
            os.remove(p)
            errs.append(f"'{fn}': 읽기 오류 {e}")
            continue
        w["files"].append(ent)
        added.append(ent["orig"])
    _save(jid, m)
    return {"added": added, "errors": errs}


@router.post("/api/jobs/{jid}/wiz/file")
def update_file(jid: str, body: dict):
    m = _m(jid)
    w = next((x for x in m["wiz"]["works"] if x["name"] == body.get("work")), None)
    if not w or not (0 <= int(body.get("idx", -1)) < len(w["files"])):
        raise HTTPException(404, "파일을 찾을 수 없습니다.")
    f = w["files"][int(body["idx"])]
    if body.get("delete"):
        p = os.path.join(base_dir(jid), f["file"])
        for x in (p, p + ".json"):
            if os.path.exists(x):
                os.remove(x)
        w["files"].pop(int(body["idx"]))
        pref = f"{w['name']}|{f['file']}|"
        m["wiz"]["items"] = {k: v for k, v in m["wiz"]["items"].items() if not k.startswith(pref)}
    else:
        if body.get("category") in CATEGORIES + [CATEGORY_AUTO]:
            f["category"] = body["category"]
        if "sheets_off" in body:
            f["sheets_off"] = list(body["sheets_off"] or [])
    _save(jid, m)
    return {"ok": True}


@router.post("/api/jobs/{jid}/wiz/state")
def set_state(jid: str, body: dict):
    """부분 저장: ov / nonwork / items(행별 병합, null=초기화) / acts / corr / corr_custom / indirect_auto / cases / step"""
    m = _m(jid)
    st = m["wiz"]
    for k, v in body.items():
        if k == "items":
            for key, o in (v or {}).items():
                if o is None:
                    st["items"].pop(key, None)
                else:
                    cur = st["items"].setdefault(key, {})
                    for kk, vv in o.items():
                        if vv in (None, ""):
                            cur.pop(kk, None)
                        else:
                            cur[kk] = vv
                    if not cur:
                        st["items"].pop(key, None)
        elif k == "acts":
            for aid, o in (v or {}).items():
                cur = st["acts"].setdefault(aid, {})
                cur.update({kk: vv for kk, vv in o.items() if vv not in (None, "")})
                for kk in [kk for kk, vv in o.items() if vv in (None, "")]:
                    cur.pop(kk, None)
        elif k == "ov":
            st["ov"].update(v or {})
            for kk in [kk for kk, vv in st["ov"].items() if vv in (None, "")]:
                st["ov"].pop(kk)
            if st["ov"].get("사업명"):
                m["name"] = st["ov"]["사업명"]
        elif k in ("nonwork", "corr", "corr_custom", "indirect_auto", "cases", "step"):
            st[k] = v
    _save(jid, m)
    return {"ok": True}


# ── 2.2 비작업조건 미리보기 ──────────────────────────────────────────────────
@router.get("/api/jobs/{jid}/wiz/nonwork")
def nonwork_preview(jid: str):
    m = _m(jid)
    st = m["wiz"]
    ov = st["ov"]
    station = ov.get("기상관측지점") or W.auto_station(ov.get("사업위치"))[0]
    nw = st.get("nonwork") or {}
    profiles = {k: {int(c): float(v) for c, v in d.items()} for k, d in (nw.get("profiles") or R.DEFAULT_PROFILES).items()}
    dust = nw.get("dust") or R.DUST_EXAMPLES[W.auto_dust(station)]
    setup = NonworkSetup(station=station, profiles=profiles, dust=dust, custom=nw.get("custom") or [0.0] * 12,
                         custom_label=nw.get("custom_label") or "사용자정의")
    try:
        start = dt.date.fromisoformat(str(ov.get("착공예정일"))[:10])
    except ValueError:
        start = dt.date.today().replace(day=1)
    rows = []
    y, mo = start.year, start.month
    for _ in range(12):
        r = {"년월": f"{y}.{mo:02d}"}
        for pn in profiles:
            x = setup.month(pn, y, mo)
            r[pn] = {"A": round(x.A, 1), "B": x.B, "C": x.C, "적용": round(x.applied, 1), "가동률": round(x.rate * 100, 1)}
        rows.append(r)
        mo += 1
        if mo > 12:
            y, mo = y + 1, 1
    annual = {pn: round(sum(r[pn]["적용"] for r in rows), 1) for pn in profiles}
    return {"station": station, "region": R.WEATHER["stations"][station]["region"], "profiles": {k: {str(c): w for c, w in v.items()} for k, v in profiles.items()},
            "dust": dust, "rows": rows, "annual": annual}


# ── 2.3 작업일수 산정(수량 ↔ 1일작업량 매칭) ─────────────────────────────────
@router.get("/api/jobs/{jid}/wiz/items")
def items(jid: str):
    m = _m(jid)
    st = m["wiz"]
    rows = W.collect(st, base_dir(jid))
    acts = W.activities(st, rows)
    corr = W.corrections(st, acts)
    out = []
    for a in acts:
        mine = [x for x in rows if x["act"] == a["공정ID"] and x["use"]]
        net = sum(x["days"] or 0 for x in mine)
        ind = sum(float(x["indirect_applied"] or 0) for x in mine)
        cor = sum(float(x["correction"] or 0) for x in mine) + sum(c[2] for c in corr if c[0] == a["공정ID"])
        out.append({"id": a["공정ID"], "name": a["공정명"], "work": a["소속공사"], "field": a["분야"], "profile": a["기상프로필"],
                    "preds": a["선행공정"], "direct": a["직접작업일수"], "overlap": a["최대중첩률(0~1)"],
                    "n": len(mine), "n_all": sum(1 for x in rows if x["act"] == a["공정ID"]),
                    "net": net, "indirect": round(ind, 1), "correction": round(cor, 1),
                    "total": round(net + ind + cor + float(a["직접작업일수"] or 0), 1),
                    "corr_items": [c[1] for c in corr if c[0] == a["공정ID"]]})
    stat = {"전체": len(rows), "반영": sum(1 for x in rows if x["use"]), "매칭불가": sum(1 for x in rows if not x["db"] and not x["rate"]),
            "자동매칭": sum(1 for x in rows if str(x["match"]).startswith(("규칙", "유사"))),
            "검토자지정": sum(1 for x in rows if x["key"] in st["items"])}
    act_choices = [{"id": a["공정ID"], "name": a["공정명"]} for a in acts]
    for w in st["works"]:
        code = R.WORK_INFO.get(w["name"], {}).get("code", "X")
        for k in W.ORDER:
            aid = f"{code}-{k}"
            if not any(c["id"] == aid for c in act_choices):
                act_choices.append({"id": aid, "name": f"[{w['name']}] {W.ACT_TPL[k][0]} (신규)"})
    return {"rows": rows, "acts": out, "stat": stat, "act_choices": act_choices,
            "corr": st.get("corr", {}), "corr_custom": st.get("corr_custom", {}), "indirect_auto": st.get("indirect_auto", True)}


@router.get("/api/jobs/{jid}/wiz/input")
def make_input(jid: str):
    """검토입력.xlsx 작성(입력서식) — 내려받기·확인용."""
    info = build(jid)
    return {"ok": True, "file": "검토입력.xlsx", "acts": len(info["acts"]), "rows": len(info["rows"])}


@router.post("/api/wiz/example")
def create_example():
    """예제로 시작: 처리시설공사(대죽 수량) + 관공사·배수설비(가상 수량) 수량파일 첨부."""
    root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "examples", "수량첨부_예제"))
    if not os.path.isdir(root):
        raise HTTPException(404, "예제 수량파일이 없습니다(tools/make_qty_example.py 실행).")
    works = [w for w in R.WORK_INFO if os.path.isdir(os.path.join(root, w))]
    jid = create({"works": works, "name": "예제 — 대죽공공폐수처리시설 고도개량 및 관로"})["id"]
    m = _m(jid)
    for w in m["wiz"]["works"]:
        dest = os.path.join(base_dir(jid), "수량", w["name"])
        os.makedirs(dest, exist_ok=True)
        for fn in sorted(os.listdir(os.path.join(root, w["name"]))):
            if fn.endswith(".xlsx"):
                shutil.copy(os.path.join(root, w["name"], fn), dest)
                w["files"].append({"file": f"수량/{w['name']}/{fn}", "orig": fn, "category": category_from_filename(fn), "sheets_off": []})
    m["wiz"]["ov"].update({"공사명": "대죽공공폐수처리시설 고도개량사업", "발주청": "서산시", "총공사비(백만원)": 9800,
                           "사업위치": "충청남도 서산시 대산읍 대죽리 일원", "착공예정일": "2027-03-01", "준공예정일": "2028-10-31",
                           "준비기간(일)": 60, "정리기간(일)": 30, "시설용량(㎥/일)": 12000})
    m["wiz"]["cases"] = [
        {"사업명": "○○공공폐수처리시설 증설사업(예시)", "시설용량": "10,000㎥/일", "공사금액(억원)": 95, "공사기간(개월)": 22, "준공년도": 2023},
        {"사업명": "△△하수처리시설 고도처리사업(예시)", "시설용량": "15,000㎥/일", "공사금액(억원)": 120, "공사기간(개월)": 24, "준공년도": 2022}]
    _save(jid, m)
    return {"id": jid}
