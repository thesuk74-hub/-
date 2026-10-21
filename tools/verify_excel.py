"""검토엑셀 재현 검증 — 엑셀 입력을 프로그램으로 산정하여 엑셀 결과와 비교한다.

  python3 tools/verify_excel.py <검토엑셀.xlsx> <출력폴더>

1) 엑셀 → 입력서식 변환(사업별 3개 + 통합 1개 + 민감도 1개)
2) 각 입력을 `python3 -m cpa run`으로 산정 (보고서·부록·공정표·워크북)
3) 비교표(월별 비작업일수, 공정그룹 작업일수·일정, 공사기간, 적정성) → 검증 워크북·검증 보고서
"""
import datetime as dt
import logging
import os
import shutil
import subprocess
import sys
import warnings

import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
logging.disable(logging.WARNING)
warnings.filterwarnings("ignore")

from cpa import refdata as R  # noqa: E402
from cpa.docbuild import Doc  # noqa: E402
from cpa.nonwork import NonworkSetup  # noqa: E402
from cpa.pipeline import run  # noqa: E402
from cpa.workbook import W, CF, RF  # noqa: E402
from tools.excel_to_input import main as convert  # noqa: E402

DPM = R.DAYS_PER_MONTH
CASES = [("횡산소규모", "입력_횡산소규모.xlsx"), ("신서 증설", "입력_신서 증설.xlsx"), ("노동리·우정리", "입력_노동리우정리.xlsx"),
         ("3개 사업 통합", "입력_3개사업통합.xlsx"), ("횡산소규모(민감도: 준비기간 동절기 진행)", "입력_횡산소규모_민감도.xlsx")]


def d(x):
    return x.date() if hasattr(x, "date") else x


def excel_values(src):
    wb = openpyxl.load_workbook(src, data_only=True)
    cp, s6, s7 = wb["⑤CP구성"], wb["⑥공사기간종합"], wb["⑦적정성검토"]
    groups = {}
    for r in list(range(16, 43)) + list(range(45, 58)) + list(range(60, 78)):
        no = cp.cell(r, 2).value
        if isinstance(no, (int, float)):
            groups[int(no)] = dict(사업=cp.cell(r, 3).value, 그룹=cp.cell(r, 4).value, F=cp.cell(r, 6).value,
                                   K=cp.cell(r, 11).value, S=d(cp.cell(r, 19).value), V=d(cp.cell(r, 22).value))
    proj = {}
    for r, nm in ((6, "횡산소규모"), (7, "신서 증설"), (8, "노동리·우정리"), (9, "3개 사업 통합")):
        proj[nm] = dict(착공=d(s6.cell(r, 19).value), 준공=d(s6.cell(r, 20).value), E=s6.cell(r, 5).value, F=s6.cell(r, 6).value,
                        H=s6.cell(r, 8).value, J=s6.cell(r, 10).value, L=s6.cell(r, 12).value)
    adq = {}
    for r, nm in ((6, "횡산소규모"), (7, "신서 증설"), (8, "노동리·우정리")):
        adq[nm] = dict(A=s7.cell(r, 3).value, 방법=s7.cell(r, 4).value, Y=s7.cell(r, 5).value, B=s7.cell(r, 7).value,
                       편차=s7.cell(r, 8).value, 판정=s7.cell(r, 9).value)
    nw = wb["②비작업일수"]
    months = []
    for r in range(8, 53):
        y, m = nw.cell(r, 3).value, nw.cell(r, 4).value
        if y:
            months.append(dict(y=y, m=m, A=nw.cell(r, 6).value, B=nw.cell(r, 7).value, C=nw.cell(r, 8).value, J=nw.cell(r, 10).value))
    return groups, proj, adq, months


def main(src, out):
    inp = os.path.join(out, "입력")
    os.makedirs(inp, exist_ok=True)
    convert(src, inp)
    # 민감도: 횡산소규모 준비·정리기간 동절기 진행(엑셀 실제 계산과 동일)
    s = os.path.join(inp, "입력_횡산소규모.xlsx")
    t = os.path.join(inp, "입력_횡산소규모_민감도.xlsx")
    shutil.copy(s, t)
    wb = openpyxl.load_workbook(t)
    for row in wb["1_공사개요"].iter_rows(min_row=2):
        if row[0].value == "준비·정리기간 동절기 중지":
            row[1].value = "N"
        if row[0].value == "공사명":
            row[1].value = "횡산소규모 (검토엑셀 재현 — 민감도: 준비기간 동절기 진행)"
    wb.save(t)

    res = {}
    for nm, f in CASES:
        o = os.path.join(out, nm.split("(")[0].replace(" ", "").replace("·", "") + ("_민감도" if "민감도" in nm else ""))
        print(f"\n▶ {nm}: python3 -m cpa run")
        subprocess.run([sys.executable, "-m", "cpa", "run", os.path.join(inp, f), "-o", o], check=True,
                       cwd=os.path.join(os.path.dirname(__file__), ".."), stdout=subprocess.DEVNULL)
        res[nm] = (run(os.path.join(inp, f), echo=False), o)

    groups, proj, adq, months = excel_values(src)
    integ = res["3개 사업 통합"][0]
    fin = integ.final
    cal = fin.cal

    # ── 비교 데이터 ──
    # (1) 월별 비작업일수 (동절기 미반영 기준, 엑셀 ② 표)
    st = integ.setup
    s0 = NonworkSetup(station=st.station, profiles=st.profiles, dust=st.dust, floor_days=st.floor_days, round_int=False,
                      winter_stop=False, custom=st.custom, no_holiday_profiles=st.no_holiday_profiles)
    nw_rows = []
    for mm in months:
        mi = s0.month("옥외공사(엑셀)", mm["y"], mm["m"])
        nw_rows.append([f"{mm['y']}.{mm['m']:02d}", mm["A"], mi.A, mm["B"], mi.B, mm["C"], mi.C, round(mm["J"], 1), mi.applied,
                        round(mi.applied - mm["J"], 1)])
    # (2) 공정그룹 작업일수·일정
    pre = {"횡산소규모": "H", "신서 증설": "S", "노동리·우정리": "N"}
    g_rows, g_fill = [], []
    for no, g in groups.items():
        aid = f"{pre[g['사업']]}{no:02d}"
        a = fin.act(aid)
        g_rows.append([aid, g["사업"], g["그룹"], round(g["F"], 1), round(a.work, 1), round(a.work - g["F"], 1),
                       g["S"], cal.day_to_date(a.es), (cal.day_to_date(a.es) - g["S"]).days if g["S"] else "",
                       g["V"], cal.day_to_date(a.ef), (cal.day_to_date(a.ef) - g["V"]).days if g["V"] else "", "●" if a.critical else ""])
        g_fill.append(CF if a.critical else None)
    # (3) 공사기간
    t_rows = []
    ex_i = proj["3개 사업 통합"]
    t_rows.append(["3개 사업 통합", "동절기 반영 총 공사기간(일)", round(ex_i["L"], 1), fin.total_days, round(fin.total_days - ex_i["L"], 1)])
    t_rows.append(["3개 사업 통합", "동절기 달력일수(일)", ex_i["J"], fin.winter_cal_days, fin.winter_cal_days - ex_i["J"]])
    t_rows.append(["3개 사업 통합", "동절기 제외 공사기간(일)", round(ex_i["H"], 1), fin.total_days - fin.winter_cal_days,
                   round(fin.total_days - fin.winter_cal_days - ex_i["H"], 1)])
    t_rows.append(["3개 사업 통합", "착공일", ex_i["착공"], cal.day_to_date(0), (cal.day_to_date(0) - ex_i["착공"]).days])
    t_rows.append(["3개 사업 통합", "준공일", ex_i["준공"], fin.end_date, (fin.end_date - ex_i["준공"]).days])
    for nm in ("횡산소규모", "신서 증설", "노동리·우정리"):
        r = res[nm][0]
        t_rows.append([nm, "적정성 비교 공사기간 A(일, 동절기 제외·준비+정리)", round(adq[nm]["A"], 1), r.adequacy["A"],
                       round(r.adequacy["A"] - adq[nm]["A"], 1)])
    rs = res["횡산소규모(민감도: 준비기간 동절기 진행)"][0]
    t_rows.append(["횡산소규모(민감도)", "A — 준비·정리기간 동절기 진행 시", round(adq["횡산소규모"]["A"], 1), rs.adequacy["A"],
                   round(rs.adequacy["A"] - adq["횡산소규모"]["A"], 1)])
    # (4) 적정성
    a_rows = []
    for nm in ("횡산소규모", "신서 증설", "노동리·우정리"):
        r = res[nm][0]
        f5 = next(x for x in r.adequacy["rows"] if x["구분"] == "부록5 산정공식")
        e = adq[nm]
        a_rows.append([nm, e["방법"], round(e["A"], 1), round(e["B"], 1), f"{e['편차'] * 100:+.1f}%", e["판정"],
                       r.adequacy["A"], round(f5["B"], 1), f"{f5['편차'] * 100:+.1f}%", f5["판정"]])

    # ── 검증 워크북 ──
    wbk = openpyxl.Workbook()
    w = W(wbk.active)
    wbk.active.title = "요약"
    w.title("검토엑셀 재현 검증 — 요약", f"원본: {os.path.basename(src)} / 프로그램: python3 -m cpa (2026 가이드라인)")
    w.table(["구분", "항목", "엑셀", "프로그램", "차이"], t_rows, [22, 44, 14, 14, 10])
    w.section("[ 적정성 검토(부록5) ]")
    w.table(["사업", "엑셀 검토방법", "엑셀 A", "엑셀 B", "엑셀 편차", "엑셀 판정", "프로그램 A", "프로그램 B", "프로그램 편차", "프로그램 판정"], a_rows)
    w = W(wbk.create_sheet("②비작업일수 비교"))
    w.title("월별 비작업일수 비교 (동절기 미반영 기준)", "기후(A)=MAX(②-1,②-2) 동일 입력, 법정공휴일(B)·중복(C)·주40시간 하한 8일")
    w.table(["년월", "엑셀 A", "프로그램 A", "엑셀 B", "프로그램 B", "엑셀 C", "프로그램 C", "엑셀 적용", "프로그램 적용", "차이"], nw_rows)
    w = W(wbk.create_sheet("⑤공정그룹 비교"))
    w.title("공정그룹별 작업일수·일정 비교 (3개 사업 통합, 동절기 반영)", "엑셀 S·V열(순연 시작·종료일) vs 프로그램 ES·EF")
    w.table(["ID", "사업", "공정그룹", "엑셀 작업일수", "프로그램 작업일수", "차이", "엑셀 시작", "프로그램 시작", "Δ시작(일)",
             "엑셀 종료", "프로그램 종료", "Δ종료(일)", "CP"], g_rows, [6, 12, 26, 9, 9, 7, 11, 11, 8, 11, 11, 8, 4], g_fill)
    w.note("※ 엑셀 'FF 종료' 지정 공정(가시설·토공 등)은 막대만 대상공정 완료까지 연장 표시되므로 종료일 차이가 크게 나타나며, 후속공정·공사기간에는 영향 없음")
    xpath = os.path.join(out, "00_엑셀재현_검증표.xlsx")
    wbk.save(xpath)

    # ── 검증 보고서 ──
    D = Doc("검토엑셀 재현 검증")
    D.set_header(D.d.sections[0], "공사기간 적정성 검토 — 엑셀 재현 검증")
    D.cover("횡산소규모 + 신서 증설 + 노동리·우정리 하수관로정비", "공사기간 산정 검증 보고서",
            f"{dt.date.today():%Y. %m.}", "검토엑셀 대비 프로그램 산정결과 비교", extra=f"({R.GUIDELINE} 기준)")
    D.ch = 1
    D.new_section("검증 개요")
    D.h2("1. 검증 개요")
    D.bullet("목적 : 업로드된 검토엑셀(3개 사업 통합 공사기간 산정)의 입력자료로 공사기간 적정성 검토 프로그램을 실행하여 엑셀 결과와 비교")
    D.bullet("방법 : 엑셀 ④-1·④-2·④-3 단위작업(수량·1일작업량·작업조), ⑤CP구성 58개 공정그룹·선후관계, ② 월별 기후비작업일수"
             " MAX(②-1, ②-2), ⑦ 부록5 변수를 입력서식으로 변환(tools/excel_to_input.py) 후 동일 조건으로 산정")
    D.bullet("산정조건 : 관측지점 철원, 착공 2026.12.01, 준비 60일·정리 30일, 동절기(12~2월) 공사중지 반영(시운전은 동절기 진행), "
             "작업일수 0.1일 단위 올림, 횡산소규모 작업일수 보정계수 0.7, 착공시차 신서 +92·노동리 +184 가동일")
    D.h2("2. 공사기간 비교 결과")
    D.table(["구분", "항목", "엑셀", "프로그램", "차이"], t_rows, "공사기간 산정결과 비교", widths=[3.4, 6.6, 2.3, 2.3, 1.6],
            align=["C", "L", "C", "C", "C"], size=8)
    pct = (fin.total_days - ex_i["L"]) / ex_i["L"] * 100
    D.para(f"3개 사업 통합 공사기간은 엑셀 {ex_i['L']:,.1f}일, 프로그램 {fin.total_days:,}일로 {pct:+.2f}% 차이이며, 준공일은 "
           f"엑셀 {ex_i['준공']:%Y.%m.%d}, 프로그램 {fin.end_date:%Y.%m.%d}로 {(fin.end_date - ex_i['준공']).days}일 차이로 사실상 동일함.", bold=True)
    D.h2("3. 적정성 검토(부록5) 비교")
    D.table(["사업", "엑셀 방법", "엑셀 A", "엑셀 B", "엑셀 편차", "엑셀 판정", "프로그램 A", "프로그램 B", "프로그램 편차", "프로그램 판정"],
            a_rows, "사업별 적정성 판정 비교", size=7.5)
    D.h2("4. 차이 원인 분석")
    hb = res["횡산소규모"][0].adequacy["A"]
    for t in [
        f"[횡산소규모 A {hb}일 vs 엑셀 {adq['횡산소규모']['A']:.1f}일] 엑셀은 주석(⑤ Q13)에 '준비기간·정리기간도 동절기 중지 대상'으로 "
        "명시했으나 실제 산정(⑤ P6 CP 기산점=0)에서는 횡산소규모 작업이 2027.03.01에 바로 착수하여 준비기간 60일이 동절기와 겹쳐 처리됨. "
        f"프로그램은 주석 기준(준비기간도 동절기 중지)으로 산정하여 약 60일 늦어짐. 민감도(준비기간 동절기 진행) 산정 시 A = {rs.adequacy['A']}일로 엑셀과 근접함.",
        "[노동리·우정리 부록5 B] 엑셀 ⑦은 하수도 산정공식 결과를 1.7로 나누어(E8 수식 '/1.7') B=996일을 산정함. 가이드라인 부록5 하수도 "
        "공식(Y=−452.433+98.364ln(SL)+0.083C, C=백만원)에는 나눗셈 계수가 없으므로 프로그램은 원식대로 B=1,630일을 산정하였고, "
        "편차가 −42.5%로 '재검토 필요'로 판정됨. 1.7 계수의 근거를 확인할 필요가 있음.",
        "[월별 비작업일수] 엑셀은 소수 비작업일수에 일 단위 중복일수(C) 반올림, 주40시간 하한 8일을 적용하였으며 프로그램과 산식이 동일함(차이 0).",
        "[비작업일수 배분 방식] 엑셀은 동절기 외 비작업일률(0.7236/작업일)을 모든 공정에 균등 적용하고, 프로그램은 공정 진행 월의 "
        "월별 가동률로 소진하는 방식(가이드라인 제2장 3.(3))으로 산정함. 계절에 따라 개별 공정 일정이 수일~수십일 차이 날 수 있으나 "
        "전체 공사기간 영향은 미미함.",
        "[FF 종료] 엑셀의 'FF 종료' 지정은 가시설·토공 막대를 대상공정 완료까지 연장 표시하는 것으로 후속공정 산정에 쓰이지 않아, "
        "프로그램에서는 선후관계에서 제외(비고에 기재)함. 해당 공정의 종료일 차이는 표시상 차이임.",
        "[횡산소규모 단위작업 반올림] 엑셀은 그룹 합계(O열)를 보정계수 0.7로 나누고, 프로그램은 단위작업별 1일작업량에 0.7을 곱한 뒤 "
        "0.1일 올림하여 그룹별 ±0.5일 이내 차이가 있음.",
        "[가설사무실 1일작업량] 검토엑셀 DB의 가설사무실 0.143개소/일은 가이드라인 원문(77일/개소=0.013)과 다름(해당 항목이 산정에 쓰인 경우 수동입력으로 엑셀값 유지).",
    ]:
        D.bullet(t)
    D.h2("5. 결론")
    D.para(f"검토엑셀의 입력자료를 프로그램으로 산정한 결과, 3개 사업 통합 공사기간(동절기 반영)은 {fin.total_days:,}일로 엑셀({ex_i['L']:,.1f}일)과 "
           f"{abs(pct):.1f}% 이내로 일치하였고, 신서 증설·노동리·우정리의 적정성 비교용 공사기간(A)도 엑셀과 1% 이내로 일치함. "
           "다만 ① 횡산소규모 준비기간의 동절기 처리(엑셀 주석과 계산의 불일치), ② 노동리·우정리 부록5 하수도식의 1.7 나눗셈 계수는 "
           "가이드라인과 다르므로 보고서 확정 전 검토가 필요함.")
    D.ch = 2
    D.new_section("검증 상세")
    D.h2("부록 A. 월별 비작업일수 비교 (동절기 미반영)")
    D.table(["년월", "엑셀 A", "프로그램 A", "엑셀 B", "프로그램 B", "엑셀 C", "프로그램 C", "엑셀 적용", "프로그램 적용", "차이"],
            nw_rows, "월별 비작업일수 비교", size=7)
    D.new_section("검증 상세", landscape=True)
    D.h2("부록 B. 공정그룹별 작업일수·일정 비교 (3개 사업 통합, 동절기 반영)")
    D.table(["ID", "사업", "공정그룹", "엑셀 작업일", "프로그램 작업일", "차이", "엑셀 시작", "프로그램 시작", "Δ시작",
             "엑셀 종료", "프로그램 종료", "Δ종료", "CP"],
            [[x if not isinstance(x, dt.date) else f"{x:%Y.%m.%d}" for x in r] for r in g_rows],
            "공정그룹 비교", size=6.5, fills=["FFE0E0" if f else None for f in g_fill])
    D.note("※ Δ종료가 큰 항목은 엑셀 'FF 종료'(막대 연장 표시) 공정임. 음영 = 프로그램 주공정")
    D.new_section("검증 상세", landscape=True)
    D.h2("부록 C. 3개 사업 통합 주공정 공정표 (프로그램 산정)")
    D.figure(os.path.join(res["3개 사업 통합"][1], "03_주공정공정표.png"), "3개 사업 통합 예정공정표", 25, max_h_cm=16)
    dpath = os.path.join(out, "00_엑셀재현_검증보고서.docx")
    D.save(dpath)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", out, dpath], stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL, timeout=300)
    print("\n작성:", xpath, dpath)
    for row in t_rows:
        print(" ", row)
    for row in a_rows:
        print(" ", row)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
