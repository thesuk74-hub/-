"""공사기간 산정(적정성 검토) 보고서 — 첨부 샘플보고서 목차·양식 준용.

제1장 사업의 개요 / 제2장 공사기간 산정(법적근거·산정·적정성 검토·결정) / 제3장 예정공정표(주공정 공정표)
"""
import datetime as dt
import os

from . import refdata as R
from .docbuild import Doc
from .summary import fmt_date, group_summary, monthly_rows

DPM = R.DAYS_PER_MONTH


def _lines(text):
    return [x.strip() for x in str(text or "").replace("\r", "").split("\n") if x.strip()]


def _d(v):
    return v.date() if isinstance(v, dt.datetime) else (v if isinstance(v, dt.date) else dt.date.fromisoformat(str(v)[:10]))


def _m(days):
    return days / DPM


def build(res, path, img_gantt, img_net, doc_cls=Doc):
    p, fin = res.p, res.final
    cal, setup = fin.cal, res.setup
    total = res.adopted_total
    adq = res.adequacy
    today = dt.date.today()
    D = doc_cls(p.name)
    D.first_header("공사기간 산정 보고서")
    D.cover(p.name, "공사기간 산정 보고서", f"{today.year}. {today.month}.", p.g("발주청", ""),
            extra=f"({R.GUIDELINE} 기준 적정성 검토)")
    D.new_section("목 차")
    toc = [(0, "제1장 사업의 개요"), (1, "1.1 사업의 개요"), (1, "1.2 사업의 주요내용"),
           (0, "제2장 공사기간 산정"), (1, "2.1 법적근거 및 산정기준"), (1, "2.2 공사기간 산정"),
           (1, "2.3 공사기간 적정성 검토"), (1, "2.4 공사기간 결정"),
           (0, "제3장 예정공정표"), (1, "3.1 주공정(Critical Path) 공정표"), (1, "3.2 시공조건 명시 및 유의사항"),
           (0, "※ 별책: 보고서 부록 (산정 기초자료)")]
    D.toc(toc)

    # ───────────── 제1장 ─────────────
    D.chapter(1, "사업의 개요", ["1.1 사업의 개요", "1.2 사업의 주요내용"])
    D.h2("1.1. 사업의 개요")
    D.h3("1.1.1 사업의 목적")
    D.para(p.g("사업목적") or f"본 사업은 {p.name}으로, 「{R.GUIDELINE}」에 따라 공사기간의 적정성을 검토하고자 함.")
    D.h3("1.1.2 사업의 개요")
    D.table(["구 분", "내 용"],
            [["공 사 명", p.name], ["발 주 청", p.g("발주청", "-")], ["사업위치", p.g("사업위치", "-")],
             ["총공사비", f"{p.total_cost:,.0f}백만원 (≒ {p.total_cost / 100:,.1f}억원)"],
             ["시설구분", p.g("시설구분", "-")], ["시설유형(부록5)", p.facility_type or "-"],
             ["착공예정일", fmt_date(p.start_date)]]
            + ([["준공예정일(공사개요)", fmt_date(_d(p.g("준공예정일")))]] if p.g("준공예정일") else [])
            + [["기상관측지점", f"{setup.station} ({setup.station_data()['region']})"]],
            "사업 개요", widths=[4, 12], align=["C", "L"], size=9.5)
    if _lines(p.g("추진경위")):
        D.h3("1.1.3 사업의 추진경위")
        for ln in _lines(p.g("추진경위")):
            D.bullet(ln, mark="◦")
    D.h2("1.2. 사업의 주요내용")
    D.h3("1.2.1 주요 시설 및 공사내용")
    lines = _lines(p.g("사업내용"))
    for ln in lines or [f"{a.field} — {a.name}" for a in p.acts]:
        D.bullet(ln, mark="◦")
    D.h3("1.2.2 분야별 공정 구성")
    gs = group_summary(fin, res.final_items)
    D.table(["분야", "대공정", "세부 공정", "작업일수(일)"],
            [[g["분야"], g["대공정"], ", ".join(g["공정"]), g["작업일수"]] for g in gs],
            "분야별 공정 구성", widths=[2.2, 3, 8.5, 2.5], align=["C", "C", "L", "C"])

    # ───────────── 제2장 ─────────────
    D.chapter(2, "공사기간 산정", ["2.1 법적근거 및 산정기준", "2.2 공사기간 산정", "2.3 공사기간 적정성 검토", "2.4 공사기간 결정"])
    D.h2("2.1. 법적근거 및 산정기준")
    for t in ["건설기술진흥법 제45조의2(공사기간 산정기준)",
              "건설기술진흥법 시행령 제17조(지방심의위원회의 구성·운영) 제2항 제3호, 제19조(기술자문위원회의 구성 및 기능 등) 제4항 제3호, 제59조(건설사업관리의 업무범위 및 업무내용) 제4항 제8호",
              f"「{R.NOTICE}」", f"「{R.GUIDELINE}」에 따라 산정"]:
        D.bullet(t)
    D.boxed(["공사기간 = ① 준비기간 + ② 비작업일수 + ③ 작업일수 + ④ 정리기간"], f"「{R.GUIDELINE}」 공사기간 산정 원칙")

    D.h2("2.2. 공사기간 산정")
    D.h3("2.2.1 금회 시설 공사특성")
    for ln in _lines(p.g("공사특성")) or ["설계내역서의 공종별 수량과 공정 선후관계를 토대로 주공정(Critical Path)을 구성함"]:
        D.bullet(ln)
    D.bullet(f"주공정(Critical Path) : {' → '.join(fin.act(i).name for i in fin.cp_ids)}")
    nc = [a.name for a in fin.acts if not a.critical and a.work > 0]
    if nc:
        D.bullet(f"주공정 외 병행작업 : {', '.join(nc)}")
    D.figure(img_net, "주공정 경로 모식도 (적색: 주공정)", 16.5)

    D.h3(f"2.2.2 「{R.GUIDELINE}」에 따른 공기산정")
    D.h4("가) 공사기간 산정 원칙")
    D.para("공사기간은 준비기간과 비작업일수, 작업일수, 정리기간을 포함하여 산정하며, 준비기간과 정리기간에 대하여는 "
           "비작업일수를 계상하지 않음. 공종별 공사기간은 해당 공종의 작업일수에 공사 개시일부터 종료일까지의 "
           "월별 비작업일수를 더하여 산정하고, 주공정(Critical Path)에 해당하는 공종별 공사기간을 합산함.")
    D.h4("나) 준비기간")
    D.para("공사 준비기간은 설계도서 검토, 안전관리계획서 작성·승인, 하도급업체 선정, 측량, 현장사무소·세륜시설·가설건물 설치, "
           "주요 건설자재·장비 및 공장제작 조달 등 공사 착수준비에 필요한 기간을 말하며, 시설물별 특성을 고려하여 반영함.")
    pp = list(R.PREP_PERIODS.items())
    half = (len(pp) + 1) // 2
    rows = [[pp[i][0], pp[i][1], pp[i + half][0] if i + half < len(pp) else "", pp[i + half][1] if i + half < len(pp) else ""] for i in range(half)]
    rows.append(["적 용", "", "", fin.prep])
    D.table(["공종", "준비기간(일)", "공종", "준비기간(일)"], rows, "공사유형별 준비기간(가이드라인 예시)", widths=[4, 3, 4, 3], bold_rows=(len(rows) - 1,))
    D.para(p.g("준비기간 산정사유") or f"금회 공사는 '{p.g('준비기간 공사유형')}' 유형의 준비기간 {fin.prep}일을 적용하였음.")
    D.note("주) 인허가 등 진행상황에 따라 준비기간이 지연될 수 있으며, 이 경우 발주청과 협의하여 재산정할 수 있음")

    D.h4("다) 비작업일수")
    D.para("비작업일수(공사불능일수)는 법정공휴일수와 기상조건으로 인한 비작업일수의 합계에서 중복일수를 제외하여 산출하며, "
           "주공정(critical path) 및 그에 영향을 미치는 공종에 대하여 공정별 기상조건(기상프로필)을 달리 적용함.")
    D.para("(1) 법정공휴일", bold=True)
    D.para("「관공서의 공휴일에 관한 규정」 제2조 및 제3조에 따른 공휴일 및 대체공휴일을 비작업일수에 포함함(가이드라인 부록1).")
    y0, y1 = cal.day_to_date(0).year, cal.day_to_date(total - 1).year
    hrows = []
    for y in range(y0, y1 + 1):
        v = [R.holidays(y, m) for m in range(1, 13)]
        hrows.append([f"{y}년"] + v + [sum(v)])
    D.table(["구분"] + [f"{m}월" for m in range(1, 13)] + ["소계"], hrows, f"월간 법정공휴일 ({y0}년~{y1}년)", size=8)
    D.para("(2) 주 40시간 근무제", bold=True)
    D.para(f"법정 근로시간(1일 8시간, 주 40시간)을 기준으로 하며, 월별 비작업일수가 주 40시간 근무제에 의한 일수(월 {setup.floor_days:g}일)보다 "
           "작을 경우 주 40시간 근무제에 따른 비작업일수를 적용함.")
    D.para("(3) 기상조건", bold=True)
    st = setup.station_data()
    D.para(f"건설공사의 주공정에 영향을 미치는 기상조건을 반영하며, 가이드라인 부록3 '기상조건별·지역별 비작업일수(2015~2024년)' 중 "
           f"{st['region']} {setup.station} 관측지점(지점번호 {st['code']})의 10년 통계를 적용함. 공종별로 작업을 제한하는 기상조건"
           "(부록2)을 검토하여 다음과 같이 기상프로필을 설정함.")
    prof_used = [x for x in setup.profiles if any(a.profile == x for a in fin.acts)]
    rows = []
    for x in prof_used:
        c = setup.profiles[x]
        txt = ", ".join(f"{R.COND_SHORT[k]}" + (f"(×{w:g})" if w != 1 else "") for k, w in sorted(c.items())) or "기상조건 미적용(연속작업)"
        appl = ", ".join(a.name for a in fin.acts if a.profile == x)
        rows.append([x, txt, appl])
    D.table(["기상프로필", "적용 기상조건 (×적용률)", "적용 공정"], rows, "작업불능 기상조건에 대한 기준 선정", widths=[2.6, 7, 6.8], align=["C", "L", "L"])
    if 14 in {k for x in prof_used for k in setup.profiles[x]}:
        D.table(["구 분", "주의보", "경보", "비상저감조치"],
                [["PM-2.5(초미세먼지)", "75㎍/㎥ 이상", "150㎍/㎥ 이상", "당일 평균 50㎍/㎥ 초과 및 익일 50㎍/㎥ 초과 예상 등"],
                 ["PM-10(미세먼지)", "150㎍/㎥ 이상", "300㎍/㎥ 이상", "-"]],
                "미세먼지 비상저감조치 발령기준(참고)", widths=[3.5, 3, 3, 7], align=["C", "C", "C", "L"])
        D.note("주) 미세먼지 비상저감조치 발령일수는 입력서식의 월별 발령일수(권역별 발령내역)를 적용함")
    for x in prof_used:
        ann = setup.annual_climate(x)
        if not ann:
            continue
        rr = [[R.COND_SHORT[k]] + v + [round(sum(v), 1)] for k, v in ann.items()]
        tot = [round(sum(v[m] for v in ann.values()), 1) for m in range(12)]
        rr.append(["총 일수"] + tot + [round(sum(tot), 1)])
        D.table(["구분"] + [f"{m}월" for m in range(1, 13)] + ["소계"], rr, f"기상조건에 의한 연간 비작업일수 — {x}",
                size=7.5, bold_rows=(len(rr) - 1,))
    D.para("(4) 월별 중복일수를 고려한 비작업일수 산정", bold=True)
    D.boxed(["비작업일수 = A + B − C",
             "A : 해당 월에 기후여건으로 인해 계획된 공종의 작업이 불가능한 일수",
             "B : 해당 월에 포함된 법정공휴일수",
             "C : 월별 중복일수 = A × B ÷ 달력일수 (소수점 첫째자리에서 반올림)",
             f"※ 월별 비작업일수 < 주40시간 근무제 비작업일수({setup.floor_days:g}일)인 경우 {setup.floor_days:g}일 적용"], fill="FFFFFF")
    mr = monthly_rows(fin, setup)
    D.para(f"착공예정일({fmt_date(p.start_date)}) 기준으로 공정별 기상프로필을 적용하여 월별 비작업일수를 산정한 결과, "
           f"주공정의 비작업일수는 {fin.cp_nonwork:.0f}일로 산정되었음. (주공정에 해당하는 공정의 기상프로필 적용, "
           "부분월은 공사 진행일수에 비례 배분)")
    for y in sorted({r["년"] for r in mr}):
        ys = [r for r in mr if r["년"] == y]
        cols = [f"{r['월']}월" for r in ys]
        f = lambda v: "-" if v is None else v  # noqa: E731
        rows = [["구분"] + [r["구분"] for r in ys] + [""],
                ["달력일수"] + [r["달력일수"] for r in ys] + [sum(r["달력일수"] for r in ys)],
                ["기후여건(A)"] + [f(r["A"]) for r in ys] + [round(sum(r["A"] or 0 for r in ys), 1)],
                ["법정공휴일(B)"] + [f(r["B"]) for r in ys] + [sum(r["B"] or 0 for r in ys)],
                ["중복일수(C)"] + [f(r["C"]) for r in ys] + [sum(r["C"] or 0 for r in ys)],
                ["비작업일수(적용)"] + [f(r["적용"]) for r in ys] + [sum(r["적용"] or 0 for r in ys)],
                ["공사기간 반영 비작업"] + [r["반영비작업"] if r["공사일수"] else "-" for r in ys] + [round(sum(r["반영비작업"] for r in ys), 1)],
                ["가동률(%)"] + [f"{r['가동률'] * 100:.1f}" if r["가동률"] is not None else "-" for r in ys] + [""]]
        D.table([str(y)] + cols + ["소계"], rows[1:], f"비작업일수 및 가동률 산정결과({y}년)", size=7.5,
                header_rows=[[f"{y}년"] + cols + ["소계"], rows[0]])
    D.note("주) 구분 : 준비·정리기간은 비작업일수 미계상, '시운전'은 연속작업(달력일) 적용. 반영 비작업 = 적용 비작업 × 공사진행일 ÷ 달력일")

    D.h4("라) 작업일수")
    D.para("(1) 산정시 고려사항", bold=True)
    D.para("작업일수는 당해 공사의 공종별 수량을 시공하는 데 필요한 총 작업일수로, 가이드라인 부록4의 공종별 1일 작업량을 활용하거나 "
           "발주청이 보유한 과거 경험치·표준품셈을 활용함. 부록4는 대표적이고 일반화된 공종 및 공법 기준이므로 현장 여건 및 공사규모, "
           "지질 조건, 기상·기후조건 등에 따라 조정하여 적용함.")
    fac = p.factors()
    t = fac.tables
    D.table(["보정요소", "적용 조건", "보정계수", "적용 대상"],
            [["현장여건", fac.site, t["현장여건"].get(fac.site, 1), "전 공종"],
             ["공사규모", fac.scale, t["공사규모"].get(fac.scale, 1), "전 공종"],
             ["지질조건", fac.geology, t["지질조건"].get(fac.geology, 1), "토공·가시설·기초 등 지반 관련 공종"],
             ["기상·기후", fac.climate, t["기상기후"].get(fac.climate, 1), "옥외·구조물 공정"]],
            "1일 작업량 보정계수 적용", widths=[3, 4, 2.5, 6.5])
    D.note("주) 적용 1일 작업량 = 기준 1일 작업량(부록4·품셈·실적) × 보정계수. 보정계수는 현장조건을 고려한 검토자 판단값임")
    D.para("(2) 작업일수 산정", bold=True)
    crit_work = sum(a.work for a in fin.acts if a.critical)
    D.para(f"내역서 수량과 1일 작업량을 매칭하여 공종별 작업일수를 산정하였으며(단위작업 작업일수 = 수량 ÷ (1일 작업량 × 작업조), "
           f"1일 미만은 1일), 동시 작업 가능여부를 고려한 주공정 작업일수는 {fin.cp_work:.0f}일로 산정되었음.")
    rows, crit_rows = [], []
    for a in sorted(fin.acts, key=lambda a: (R.FIELDS.index(a.field) if a.field in R.FIELDS else 99, a.es)):
        its = res.final_items.get(a.id, [])
        appl = ", ".join(dict.fromkeys(it.name for it in its))[:80] or (f"직접입력 {a.direct_days:g}일" if a.direct_days else "-")
        rows.append([a.field, a.group, a.name, appl, a.work, "●" if a.critical else ""])
        if a.critical:
            crit_rows.append(len(rows) - 1)
    rows.append(["합 계", "", "", "단순합계 / 주공정 작업일수 합계", sum(a.work for a in fin.acts), f"CP {crit_work:.0f}일"])
    D.table(["분야", "대공정", "공정", "적용항목", "작업일수", "CP"], rows, "작업일수 산정 (단위 : 일)",
            widths=[1.6, 2, 3.6, 6.6, 1.6, 1.4], align=["C", "C", "L", "L", "C", "C"], size=8,
            fills=[("FFF2F2" if i in crit_rows else None) for i in range(len(rows))], bold_rows=(len(rows) - 1,))
    D.note(f"주) 1. 가이드라인 부록4·표준품셈·실적자료 적용, 세부 산정내역은 [보고서 부록] 참조  2. 음영 = 주공정")
    detail = []
    for a in fin.acts:
        for it in res.final_items.get(a.id, []):
            detail.append([a.group, f"{it.name} ({it.spec})" if it.spec else it.name, it.unit, it.qty,
                           round(it.rate, 2) if it.rate else "-", it.crews, "중복가능" if it.parallel else it.days])
    if detail:
        D.table(["대공정", "공 종", "단위", "수량", "적용 1일작업량", "작업조", "작업일수(일)"], detail,
                "작업일수 산정 세부공종별", widths=[2, 7, 1.4, 1.8, 2, 1.3, 1.8], align=["C", "L", "C", "R", "R", "C", "C"], size=7.5)
        D.note("주) 1일 미만은 1일로 표기, '중복가능'은 타 작업과 동시 진행되어 작업일수에 합산하지 않음")

    D.h4("마) 정리기간 산정")
    D.para(f"정리기간은 공정상 여유기간(buffer)과는 다르며 준공검사 준비, 시설물 인수 등을 위한 행정절차 및 청소 등 현장정리에 필요한 기간으로, "
           f"주요공종이 마무리된 이후 준공 전 1개월의 범위에서 계상할 수 있음. 금회 정리기간은 {fin.cleanup}일로 적용하였음.")
    D.h4("바) 공사여건 등에 따른 보정")
    if res.correction_days:
        D.para(f"공사의 규모 및 성격, 지역여건 등을 고려하여 {res.correction_days}일을 추가 반영하였음. {p.g('공사여건 보정사유', '')}")
    else:
        D.para("공사의 규모 및 성격(신설/확장/개량 등), 지역여건(산지/농경지/도심지/도서지역) 등을 고려하여 추가 공사기간을 반영할 수 있으나, "
               "금회 공사기간 산정 시에는 별도의 추가 공사기간을 반영하지 않았음.")
    D.h4("사) 기타 공사기간 지연사항")
    D.para(p.g("시공조건 및 지연요인") or "인허가 진행상황, 지장물 이설 및 철거, 문화재 시·발굴, 지반조사 내용 및 관계기관 협의 등으로 공사 계약기간의 "
           "지연이 발생할 수 있으며, 이에 따른 기간은 별도 산정하여야 하므로 공사기간에 반영하지 않았음.")

    D.h4("아) 주공정(Critical Path) 구성")
    D.para("공정별 작업일수와 선후관계(FS: 완료 후 착수, SS: 선행공정 일정 진행 후 착수, FF: 동시 완료)를 적용하여 CPM 전진·후진 계산을 수행하고, "
           "총여유(TF)가 0인 공정을 주공정으로 선정하였음. 공종별 공사기간은 작업일수에 해당 기간의 비작업일수를 더하여 산정함.")
    rows = [["준비기간", "-", "-", "-", fmt_date(cal.day_to_date(0)), fmt_date(cal.day_to_date(fin.prep_end)), fin.prep]]
    for i in fin.cp_ids:
        a = fin.act(i)
        rows.append([a.name, ", ".join(r.label() for r in a.preds) or "-", round(a.work), a.cal_days - round(a.work),
                     fmt_date(cal.day_to_date(a.es)), fmt_date(cal.day_to_date(a.ef)), a.cal_days])
    rows.append(["정리기간", "-", "-", "-", fmt_date(cal.day_to_date(fin.finish_day + 1)), fmt_date(fin.end_date), fin.cleanup])
    D.table(["주공정 공종", "선행관계", "작업일수", "비작업일수", "착수", "완료", "공종별 공사기간"], rows,
            "주공정 공종별 공사기간", widths=[4.5, 2.6, 1.6, 1.6, 2.2, 2.2, 2], align=["L", "C", "C", "C", "C", "C", "C"], size=8)
    D.note("주) 선행공정과 중첩(SS·FF)되는 구간은 공종별 공사기간의 단순합계에서 제외되어 전체 공사기간이 산정됨")

    D.h4("자) 산정 결과")
    D.para(f"「{R.NOTICE}」 및 「{R.GUIDELINE}」에 따라 공사기간을 산정한 결과, 총 {total:,}일(약 {_m(total):.1f}개월)로 산정되었음.")
    tol = adq["tol"]
    rows = [["합 계", "공사기간", total, f"≒{_m(total):.1f}", f"−{tol * 100:.0f}% {_m(total) * (1 - tol):.1f}개월 / +{tol * 100:.0f}% {_m(total) * (1 + tol):.1f}개월"],
            ["1", "준비기간", fin.prep, round(_m(fin.prep), 1), "제7조 (준비기간 산정)"],
            ["2", "비작업일수", round(fin.cp_nonwork), round(_m(fin.cp_nonwork), 1), "제8조~제10조 (비작업일수)"],
            ["3", "작업일수", round(fin.cp_work), round(_m(fin.cp_work), 1), "제11조 (작업일수)"],
            ["4", "정리기간", fin.cleanup, round(_m(fin.cleanup), 1), "제12조 (정리기간 산정)"]]
    if fin.winter_stop_days:
        rows.append(["5", "동절기 공사중지", round(fin.winter_stop_days), round(_m(fin.winter_stop_days), 1),
                     f"{','.join(map(str, p.winter_months))}월 공사중지(준비·정리기간 포함)"])
    if res.correction_days:
        rows.append(["6", "공사여건 보정", res.correction_days, round(_m(res.correction_days), 1), "제2장 2.(6)"])
    D.table(["연번", "구분", "일수", "개월", "기준"], rows, "공공 건설공사의 공사기간 산정기준에 따른 공사기간 산정 결과",
            widths=[1.6, 3, 2, 2, 7.5], bold_rows=(0,))
    D.para(f"착공예정일 {fmt_date(cal.day_to_date(0))} 기준 준공예정일은 {fmt_date(cal.day_to_date(total - 1))}임.")

    # 2.2.3 공기조정
    if res.target_days:
        D.h3("2.2.3 목표기간 대비 공사기간 조정 검토")
        b = res.base
        D.para(f"발주청 요구 등 목표기간({res.target_desc}, {res.target_days}일)과 산정 공사기간을 비교하여, 초과 시 작업조 증원 및 "
               "동시 수행이 가능한 공정의 일정 중첩(Fast-tracking)을 통해 목표기간 내 작업일수를 조정하였음.")
        if b.total_days <= res.target_days:
            D.para(f"기본 산정 공사기간 {b.total_days}일은 목표기간 {res.target_days}일 이내로, 공기조정이 필요하지 않음.")
        else:
            D.para(f"기본 산정 공사기간 {b.total_days}일({_m(b.total_days):.1f}개월)이 목표기간을 {b.total_days - res.target_days}일 초과하여 "
                   f"다음과 같이 조정한 결과 {fin.total_days}일({_m(fin.total_days):.1f}개월)로 "
                   + ("목표기간을 충족함." if fin.total_days <= res.target_days else "목표기간을 충족하지 못하였음. 목표기간의 재검토 또는 공법·작업시간(교대근무 등) 변경 검토가 필요함."))
            if res.compress_history:
                D.table(["단계", "구분", "조치내용", "조정 전(일)", "조정 후(일)"],
                        [[h["단계"], h["구분"], h["조치내용"], h["조정전(일)"], h["조정후(일)"]] for h in res.compress_history],
                        "공기조정 조치내역", widths=[1.2, 1.6, 9.6, 2, 2], align=["C", "C", "L", "C", "C"], size=8)
            D.note("주) 작업조 증원 시 법정 근로시간(1일 8시간, 주 40시간) 준수를 전제로 하며, 동시작업에 따른 안전·품질 관리계획 수립 필요")

    # 2.3 적정성 검토
    D.h2("2.3. 공사기간 적정성 검토")
    D.para(f"가이드라인 제3장에 따라 산정 공사기간(A)을 시설물별 공사기간 산정공식(부록5) 또는 최근 5년간 준공된 동종공사의 실제 "
           f"공사기간 평균에 준비기간과 정리기간을 합산한 공사기간(B)과 비교하여, ±{tol * 100:.0f}% 범위 이내인 경우 적정한 것으로 판단함.")
    for r in adq["rows"]:
        if r["구분"] == "부록5 산정공식":
            D.h3("2.3.1 시설물별 공사기간 산정공식에 의한 공사기간 (가이드라인 부록5)")
            D.boxed([f"{r['기준']} : {r['산정식']}", f"({r['변수']})", "☞ 산정공식에 의해 산정된 값(Y)은 비작업일수를 포함하며, 준비기간 및 정리기간을 합산함"])
            if r.get("Y"):
                vv = {**p.formula_vars, "C(억원)": p.total_cost / 100, "C(백만원)": p.total_cost}
                D.table(list(vv.keys()) + ["Y(일)"], [[round(v, 2) for v in vv.values()] + [round(r["Y"])]], "산정공식 변수 및 산정값", size=8.5)
                D.table(["연번", "구분", "일수", "개월", "비고"],
                        [["합계", "공사기간(B)", round(r["B"]), round(_m(r["B"]), 1), r.get("비고", "")],
                         ["1", "준비기간", fin.prep, round(_m(fin.prep), 1), ""], ["2", "산정공식(Y)", round(r["Y"]), round(_m(r["Y"]), 1), "비작업일수 포함"],
                         ["3", "정리기간", fin.cleanup, round(_m(fin.cleanup), 1), ""]], "산정공식에 의한 공사기간 산정 결과",
                        widths=[1.6, 3.2, 2, 2, 7], bold_rows=(0,))
            else:
                D.para(r.get("비고", ""))
        elif r["구분"] == "동종공사 실적":
            D.h3("2.3.2 최근 준공된 동종·유사공사의 공사기간")
            cs = p.cases
            D.table(["사업명", "시설용량", "공사금액(억원)", "공사기간(개월)"],
                    [[c["사업명"], c.get("시설용량") or "-", c.get("공사금액(억원)") or "-", c.get("공사기간(개월)")] for c in cs]
                    + [["평 균", "-", "", round(sum(c["공사기간(개월)"] for c in cs if c.get("공사기간(개월)")) / len([c for c in cs if c.get("공사기간(개월)")]), 1)]],
                    "최근 준공된 동종·유사공사의 공사기간", widths=[7.5, 3, 3, 3], align=["L", "C", "C", "C"], bold_rows=(len(cs),))
            D.para(f"동종·유사공사 실적 평균 {_m(r['Y']):.1f}개월" + (" 에 준비·정리기간을 합산한" if p.cases_add_prep else "을 적용한") +
                   f" 공사기간(B)은 {round(r['B'])}일({_m(r['B']):.1f}개월)임.")
        elif r["구분"] == "건설사업관리기간":
            D.h3("2.3.3 건설사업관리기간(공사비 기준)에 의한 공사기간 (참고)")
            D.table(["공사비(억원)"] + [str(x) for x, _ in R.CM_PERIOD_TABLE], [["관리기간(개월)"] + [y for _, y in R.CM_PERIOD_TABLE]],
                    "공사비에 따른 평균 건설사업관리기간", size=7.5)
            D.para(f"{R.CM_SOURCE}에 따라 공사비 {p.total_cost / 100:,.1f}억원을 직선보간한 평균 건설사업관리기간은 {_m(r['B']):.1f}개월임.")
        elif r["구분"].startswith("환경공단"):
            D.h3(f"2.3.{4 if r['구분'] == '환경공단 산정식' else 5} {r['구분']} (참고)")
            D.para(f"{r['기준']} : {r['산정식']}" + (f" → Y = {r['Y']:.0f}일, B = Y + 준비 + 정리 = {r['B']:.0f}일" if r.get("Y") else f" → {_m(r['B']):.0f}개월"))
    subs = __import__("cpa.service", fromlist=["sub_rows"]).sub_rows(res)
    if subs:
        D.h3("2.3.5 복합공사 공사별 공사기간 및 적정성")
        D.para("복합공사는 공사별 입력(수량·공정)으로 각각 독립 산정하여 공사별 적정성을 검토하고, 통합 공정망으로 전체 공사기간을 산정함.")
        D.table(["공사명", "시설유형", "공사비(백만원)", "통합 착수", "통합 완료", "공사기간(일)", "A(일)", "B(일)", "편차", "판정"],
                [[x["공사명"], x["시설유형"], x["공사비_백만원"], x["통합_착수"], x["통합_완료"], x["독립_공사기간_일"], x["A_일"],
                  x["B_일"] if x["B_일"] else "-", f"{x['편차_%']:+.1f}%" if x["편차_%"] is not None else "-", x["판정"]] for x in subs],
                "공사별 공사기간 및 적정성 판정", size=7.5)
        D.note("주) 공사기간·A·B는 공사별 독립 산정값(각 공사 준비·정리기간 포함), 통합 착수·완료는 통합 공정망에서의 해당 공사 공정 기간. 공사별 상세는 별도 공사별 보고서 참조")
    D.h3("2.3.6 적정성 판정")
    rows, fills = [], []
    for r in adq["rows"]:
        if not r.get("B"):
            continue
        rows.append([r["구분"], r["기준"][:40], round(r["B"]), round(_m(r["B"]), 1), adq["A"], f"{r['편차'] * 100:+.1f}%", r["판정"]])
        fills.append("E2EFDA" if r.get("주기준") else None)
    D.table(["검토기준", "세부기준", "B(일)", "B(개월)", "A(일)", "편차 (A−B)/B", "판정"], rows,
            f"공사기간 적정성 판정 (허용범위 ±{tol * 100:.0f}%)", widths=[2.8, 5.6, 1.5, 1.5, 1.5, 2, 2],
            align=["C", "L", "C", "C", "C", "C", "C"], size=8, fills=fills)
    D.note("주) 음영 = 가이드라인 적정성 판정 기준(부록5 적용범위 내 산정공식, 동종공사 실적), 그 외는 참고")
    if adq.get("winter_excluded"):
        D.para(f"동절기 공사중지를 반영한 공사기간 {total:,}일에서 공사기간 내 동절기 달력일수 {adq['winter_excluded']}일을 제외한 "
               f"{adq['A']:,}일을 비교용 공사기간(A)으로 적용함(부록5 산정공식·실적공기는 동절기 공사중지를 포함하지 않으므로 산정기준을 일치시킴).")
    D.para(f"검토 결과 산정 공사기간(A) {adq['A']:,}일은 [{adq['basis']}] 기준 대비 ±{tol * 100:.0f}% 범위에서 '{adq['verdict']}'으로 판단됨.",
           bold=True)
    if "재검토" in adq["verdict"] or "일부" in adq["verdict"]:
        D.para("편차가 허용범위를 벗어난 기준에 대해서는 시설물 특성(산정공식의 대상 시설과의 차이), 공사비 구성, 동종사례의 규모 차이 등을 "
               "검토하였으며, 필요 시 전문가 자문을 통해 공사기간(A)의 산정과정을 재검토할 필요가 있음.")

    # 2.4 결정
    D.h2("2.4. 공사기간 결정")
    D.para(f"「{R.GUIDELINE}」에 따라 공사 현장의 여건을 반영한 설계내역서 기반의 공정별 작업일수, 법정공휴일 및 지역 기상조건을 반영한 "
           f"비작업일수, 주공정(CPM) 분석을 통해 산정된 {total:,}일(약 {_m(total):.1f}개월)을 공사기간으로 결정함.")
    rows = [["1", "공공 건설공사의 공사기간 산정기준(가이드라인)", f"{_m(total):.1f}개월", "채 택"]]
    for k, r in enumerate([r for r in adq["rows"] if r.get("B")], 2):
        rows.append([str(k), f"{r['구분']} — {r['기준'][:30]}", f"{_m(r['B']):.1f}개월", "비교"])
    D.table(["구 분", "산 정 기 준", "기 간", "비 고"], rows, "공사기간 결정", widths=[1.5, 10, 2.5, 2], align=["C", "L", "C", "C"], bold_rows=(0,))
    if res.winter:
        wi = res.winter
        D.note(f"주) 동절기({','.join(map(str, p.winter_months))}월) 공사중지를 반영할 경우 공사기간은 {wi.total_days:,}일({wi.months:.1f}개월), "
               f"준공예정일 {fmt_date(wi.end_date)}로 {wi.total_days - fin.total_days}일 증가하므로, 발주청은 시공조건 명시사항에 동절기 공사중지 "
               "여부를 명확히 기재하고 계약공기 산정 시 반영하여야 함.", indent=0.3)
    D.note("주) 준비기간 및 정리기간은 인허가 진행상황 및 동절기 공사휴무 조건에 따라 조정될 수 있음", indent=0.3)

    # ───────────── 제3장 ─────────────
    D.chapter(3, "예정공정표", ["3.1 주공정(Critical Path) 공정표", "3.2 시공조건 명시 및 유의사항"], title_page=False)
    from PIL import Image
    gw, gh = Image.open(img_gantt).size
    land = gh / gw < 1.0
    D.new_section("제 3 장 예정공정표", landscape=land)
    D.h2("3.1. 주공정(Critical Path) 공정표")
    D.figure(img_gantt, f"{p.name} 예정공정표", 20.5 if land else 16.5, max_h_cm=16.5 if land else 23.5)
    D.new_section("제 3 장 예정공정표")
    D.h2("3.2. 시공조건 명시 및 유의사항")
    D.para("가이드라인 부록6에 따라 공사기간에 영향을 미칠 수 있는 다음 사항의 진행상황을 입찰참가자에게 제공하여야 함.")
    for t in ["용지보상 : 미처리 부분의 장소·범위 및 처리 예정시기, 공사용지 사용종료 후 복구내용",
              "인·허가 및 심의 절차 : 관계기관 협의 미완료 사항, 시공시기·시간·방법의 제한사항",
              "지장물 이설·철거 : 지하매설물 사전조사 기간, 지장물 관리자·위치·이설시기",
              "문화재 시·발굴 : 사전조사 및 시·발굴 예정기간",
              "환경 관련 : 소음·진동·분진 대책, 탁수·용수 처리, 지반침하 우려 시 사전·사후 조사",
              "안전대책 : 근접시설 공사의 작업시간 제한, 교통유도원·보안요원 배치",
              "동절기 공사중지 여부 및 기간 (추가검토 결과 참조)"]:
        D.bullet(t)
    D.para("본 보고서의 산정 기초자료(관측지점 기상자료, 1일 작업량 매칭, 작업일수 산정 세부내역, CPM 계산표, 공기조정 이력 등)는 "
           "별책 [보고서 부록]에 수록하였음.", italic=True)
    D.save(path)
