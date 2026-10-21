"""검토엑셀(3개 사업 통합: 횡산소규모·신서 증설·노동리·우정리) → 프로그램 입력서식 변환.

엑셀의 산정값을 그대로 재현할 수 있도록 다음을 옮긴다.
 · ④-1/④-2/④-3 단위작업(수량·적용1일작업량·작업조·간접·보정) — 1일작업량이 부록4 DB와 같으면 DB지정, 다르면 수동입력
 · ⑤CP구성 58개 공정그룹과 선후관계(START/SS 진행률/FS/FS-/FF 종료/시운전 통수조건)
 · ② 월별 기후비작업일수 = MAX(②-1, ②-2) → 사용자정의 기상조건(15)으로 입력
 · ⑦ 부록5 변수(사업별 총공사비·연면적·층수·하수도연장)
출력: 사업별 3개 + 통합 1개 입력서식
"""
import datetime as dt
import os
import re
import sys
import warnings

import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from cpa import refdata as R  # noqa: E402
from cpa.project import write_template  # noqa: E402

warnings.filterwarnings("ignore")

PROJ = {  # 사업, 접두, ⑤ 행 범위, 작업일수 보정계수(F = O ÷ 계수), 비고
    "횡산소규모": ("H", range(16, 43), 0.7),
    "신서 증설": ("S", range(45, 58), 1.0),
    "노동리·우정리": ("N", range(60, 78), 1.0),
}
OFFSET = {"신서 증설": 92, "노동리·우정리": 184}   # ⑤ P열 CP 기산점 누적가동일 (동절기 제외 축)


def field_of(name):
    for k, f in (("건축", "건축"), ("기계", "기계"), ("전기", "전기"), ("조경", "조경"), ("시운전", "시운전")):
        if name.startswith(k) or (k == "시운전" and "시운전" in name):
            return f
    return "토목"


def main(src, outdir):
    wv = openpyxl.load_workbook(src, data_only=True)
    wf = openpyxl.load_workbook(src)
    cp_v, cp_f = wv["⑤CP구성"], wf["⑤CP구성"]
    db_by_key = {(it["group"], it["name"]): it for it in R.PRODUCTIVITY["items"]}

    # ② 월별 기후비작업일수 MAX(②-1, ②-2)
    a1 = [wv["②-1 기후비작업근거"].cell(69, c).value or 0 for c in range(6, 18)]
    a2 = [wv["②-2 기후비작업근거"].cell(72, c).value or 0 for c in range(6, 18)]
    climate = [round(max(x, y), 2) for x, y in zip(a1, a2)]

    proj_items, proj_acts = {}, {}
    num2id = {}
    for pname, (pre, rows, coef) in PROJ.items():
        for r in rows:
            no = cp_v.cell(r, 2).value
            if isinstance(no, (int, float)):
                num2id[int(no)] = f"{pre}{int(no):02d}"
    for pname, (pre, rows, coef) in PROJ.items():
        items, acts = [], []
        for r in rows:
            no = cp_v.cell(r, 2).value
            if not isinstance(no, (int, float)):
                continue
            no = int(no)
            aid = num2id[no]
            grp = cp_v.cell(r, 4).value
            ref = cp_v.cell(r, 5).value            # '④-2 횡산소규모 작업일수!O5'
            F = cp_v.cell(r, 6).value
            H = cp_v.cell(r, 8).value or 0
            typ, pred, ffq, lag, wint = (cp_v.cell(r, c).value for c in (27, 28, 30, 31, 32))
            fld = field_of(grp)
            profile = "옥외공사(엑셀)" if fld == "토목" else ("시운전" if fld == "시운전" else "연속작업")
            preds = []
            if typ == "SS" and pred:
                frac = lag if isinstance(lag, (int, float)) else 1 - H   # 엑셀: IF(AE="", 1-H, AE) — 0도 유효
                preds.append(num2id[int(pred)] if frac >= 0.999 else f"{num2id[int(pred)]}:SS{frac * 100:.1f}%")
            elif typ == "FS" and pred:
                preds.append(num2id[int(pred)])
            elif typ == "FS-" and pred:
                preds.append(f"{num2id[int(pred)]}:FS-{int(lag)}")
            elif typ == "시운전":
                # 횡산: 통수조건 MAX(처리시설 토목 완료, 오수관로 완료, 배수설비 50%, 기계·전기 완료) / 신서: 건축·기계·전기 완료 후
                preds += ([num2id[7], num2id[13], num2id[24], num2id[25]] if pre == "H" else [num2id[36], num2id[37], num2id[38]])
            # 'FF 종료' = 막대를 대상 공정 완료까지 연장 표시(해먹). 후속공정 산정에 쓰이지 않으므로 선후관계에서 제외
            tag = {"H": "[횡산]", "S": "[신서]", "N": "[노동리]"}[pre]
            act = {"공정ID": aid, "분야": fld, "대공정": pname, "공정명": f"{tag} {grp}", "기상프로필": profile,
                   "선행공정": ", ".join(preds), "비고": f"엑셀 ⑤ 순번 {no}, 관계 {typ}" + (f", 대상공정 {num2id[int(ffq)]} 완료까지 유지(FF 종료, 일정 영향 없음)" if ffq else "")}
            if ref:
                sh, cell = ref.split("!")
                ws_v, ws_f = wv[sh], wf[sh]
                hdr_row = int(re.sub(r"\D", "", cell))
                m = re.match(r"=SUM\(O(\d+):O(\d+)\)", str(ws_f.cell(hdr_row, 15).value))
                a, b = int(m.group(1)), int(m.group(2))
                for rr in range(a, b + 1):
                    q = ws_v.cell(rr, 7).value
                    if not isinstance(q, (int, float)) or q == 0:
                        continue
                    c_grp, c_name, c_sub, src_ = (ws_v.cell(rr, c).value for c in (3, 4, 5, 6))
                    unit = str(ws_v.cell(rr, 8).value or "").replace("/일", "")
                    J = ws_v.cell(rr, 10).value
                    K = ws_v.cell(rr, 11).value or 1
                    M_ = ws_v.cell(rr, 13).value or 0
                    N_ = ws_v.cell(rr, 14).value or 0
                    db = db_by_key.get((c_grp, c_name))
                    row = {"No": len(items) + 1, "공정ID": aid, "공종명(내역)": c_name, "규격": c_sub or "", "단위": unit,
                           "수량": q, "작업조": K, "간접작업일수": M_ or None, "보정일수": N_ or None,
                           "개별보정계수": coef if coef != 1 else None}
                    if db and J and abs(db["rate"] - J) < 1e-6:
                        row["DB지정(번호/공종명)"] = db["no"]
                        row["비고"] = f"{sh} {rr}행, 부록4 DB"
                    else:
                        row["수동1일작업량(품셈·실적)"] = J
                        why = "엑셀 적용값" + (f"(DB {db['display']}와 상이)" if db else "")
                        row["비고"] = f"{sh} {rr}행, {src_ or '품셈·실적'} {why}"
                    items.append(row)
                if coef != 1:
                    act["비고"] += f", 작업일수 보정계수 {coef}(엑셀 Z26) → 1일작업량×{coef}"
            else:
                act["직접작업일수"] = F
            acts.append(act)
        proj_items[pname], proj_acts[pname] = items, acts

    ov7 = wv["⑦적정성검토"]
    pvars = {
        "횡산소규모": dict(cost=ov7["C14"].value * 100, A=ov7["D14"].value, G=ov7["E14"].value, B=ov7["F14"].value, ftype="기타건축물"),
        "신서 증설": dict(cost=ov7["C15"].value * 100, A=ov7["D15"].value, G=ov7["E15"].value, B=ov7["F15"].value, ftype="기타건축물"),
        "노동리·우정리": dict(cost=ov7["C16"].value * 100, SL=ov7["D16"].value, ftype="하수도"),
    }
    ov1 = wv["①공사개요"]
    start = ov1["C10"].value.date()
    profiles = {"옥외공사(엑셀)": {15: 1.0}, "연속작업": {}, "시운전": {}}
    common = {
        "발주청": ov1["C6"].value, "사업위치": "경기도 연천군 일원", "기상관측지점": ov1["C11"].value,
        "준비기간 공사유형": "상수도공사", "준비기간(일)": 60, "정리기간(일)": 30,
        "비작업일수 정수반올림": "N", "주40시간 하한(일/월)": 8, "동절기 공사중지": "반영", "동절기 중지월": "12,1,2",
        "현장여건": "양호", "공사규모": "보통", "지질조건": "양호(토사)", "기상기후": "일반", "작업일수 산정단위(일)": 0.1,
        "목표기간 기준": "없음", "유사사례 준비·정리 가산": "Y",
        "준비기간 산정사유": "검토엑셀과 동일하게 상수도공사 준비기간 예시(60일)를 적용",
        "시공조건 및 지연요인": "사업대상지(연천군)는 접경·한랭지로 동절기(12~2월) 공사중지를 반영하며, 준비·정리기간도 동절기 중지 대상으로 처리함(검토엑셀 기준).",
    }
    common["준비·정리기간 동절기 중지"] = "Y"
    kw = dict(custom=climate, custom_label="엑셀 월별 기후비작업일수 MAX(②-1,②-2)", holiday_off={"연속작업", "시운전"},
              winter_exempt={"시운전"})
    os.makedirs(outdir, exist_ok=True)
    outs = []
    # 사업별(독립) — 엑셀 ⑦ 적정성 검토 비교용: 각 사업 작업착수 시점(동절기 제외 축 기산점)을 맞춤
    work_start = {"횡산소규모": None, "신서 증설": dt.date(2027, 6, 1), "노동리·우정리": dt.date(2027, 9, 1)}
    for pname in PROJ:
        v = pvars[pname]
        ov = dict(common)
        ov.update({"공사명": f"{pname} (검토엑셀 재현)", "시설유형(부록5)": v["ftype"], "총공사비(백만원)": v["cost"],
                   "착공예정일": start if work_start[pname] is None else work_start[pname] - dt.timedelta(days=60)})
        for k, key in (("연면적(100㎡)", "A"), ("지상층수", "G"), ("지하층수", "B"), ("하수도연장(m)", "SL")):
            if v.get(key) is not None:
                ov[k] = v[key]
        ov["공사특성"] = f"검토엑셀 ⑤CP구성의 {pname} 공정그룹·선후관계를 그대로 적용\n작업일수: 엑셀 ④ 단위작업 수량·1일작업량·작업조(0.1일 단위 올림)"
        path = os.path.join(outdir, f"입력_{pname.replace('·', '')}.xlsx")
        write_template(path, ov, profiles, [0] * 12, proj_items[pname], proj_acts[pname], [], **kw)
        outs.append(path)
    # 통합(착공 시차 반영) — 엑셀 ⑥공사기간종합·추가검토 비교용
    items, acts = [], []
    for pname in PROJ:
        for it in proj_items[pname]:
            it = dict(it)
            it["No"] = len(items) + 1
            items.append(it)
        for a in proj_acts[pname]:
            a = dict(a)
            if pname in OFFSET and not a["선행공정"]:
                a["선행공정"] = f"X{PROJ[pname][0]}"
            acts.append(a)
    for pname, d in OFFSET.items():
        acts.append({"공정ID": f"X{PROJ[pname][0]}", "분야": "기타", "대공정": "착공시차", "공정명": f"{pname} 착공 시차",
                     "기상프로필": "연속작업", "직접작업일수": d - 60, "선행공정": "",
                     "비고": f"엑셀 ⑤ CP 기산점 누적가동일 {d}일(준비 60일 포함) — 동절기 제외 가동일 기준"})
    ov = dict(common)
    ov.update({"공사명": "횡산소규모 + 신서 증설 + 노동리·우정리 하수관로정비 (3개 사업 통합, 검토엑셀 재현)",
               "시설유형(부록5)": "해당없음", "총공사비(백만원)": sum(v["cost"] for v in pvars.values()), "착공예정일": start,
               "공사특성": "3개 사업 병행 시공(착공 시차: 신서 +92 가동일, 노동리·우정리 +184 가동일)\n준비기간은 최초 착공사업에, 정리기간은 최종 준공사업에 1회 계상"})
    path = os.path.join(outdir, "입력_3개사업통합.xlsx")
    write_template(path, ov, profiles, [0] * 12, items, acts, [], **kw)
    outs.append(path)
    print("기후비작업일수 MAX(②-1,②-2):", climate)
    for o in outs:
        print("작성:", o)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
