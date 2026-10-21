"""보고서 부록 — 공사기간 산정 기초자료(첨부)."""
import datetime as dt

from . import refdata as R
from .docbuild import Doc
from .summary import fmt_date, monthly_rows

DPM = R.DAYS_PER_MONTH

SECTIONS = ["부록 1. 법정 공휴일수", "부록 2. 작업제한 기상조건 및 기상프로필 설정", "부록 3. 관측지점 기상조건별 비작업일수 및 월별 산정표",
            "부록 4. 1일 작업량 적용내역", "부록 5. 작업일수 산정 세부내역", "부록 6. CPM 계산표 (주공정 분석)",
            "부록 7. 공사기간 조정 이력", "부록 8. 적정성 검토 산정근거", "부록 9. 동절기 공사중지 추가검토",
            "부록 10. 시공조건 명시 항목 예시 (가이드라인 부록6)", "부록 11. 검토 작업경과"]


def build(res, path):
    p, fin = res.p, res.final
    cal, setup = fin.cal, res.setup
    D = Doc(p.name)
    D.set_header(D.d.sections[0], "보고서 부록")
    today = dt.date.today()
    D.cover(p.name, "공사기간 산정 보고서 부록", f"{today.year}. {today.month}.", p.g("발주청", ""), extra="(공사기간 산정 기초자료)")
    D.new_section("부록 목차")
    D.toc([(0, s) for s in SECTIONS])
    D.ch = "부록"

    # 부록1
    D.new_section(SECTIONS[0])
    D.h2(SECTIONS[0])
    D.table(["연도"] + [f"{m}월" for m in range(1, 13)] + ["소계"],
            [[f"{y}년"] + v + [sum(v)] for y, v in R.HOLIDAYS["data"].items()], "법정 공휴일수(2026~2035년)", size=8)
    D.note("※ 「관공서의 공휴일에 관한 규정」에 따른 공휴일·대체공휴일. 임기만료 선거일 및 수시 지정일은 불규칙하여 제외 (가이드라인 부록1)")

    # 부록2
    D.new_section(SECTIONS[1])
    D.h2(SECTIONS[1])
    D.table(["공종", "강우", "혹서기", "동절기", "신적설", "바람", "미세먼지"],
            [["토공사, 가시설공사 등(옥외공사)", "5㎜ 이상", "체감 33℃ 이상", "0℃ 이하", "5㎝ 이상", "15m/s 이상", "PM2.5 보정"],
             ["구조물공사(콘크리트 타설)", "3㎜ 이상", "체감 33℃ 이상", "0℃ 이하", "5㎝ 이상", "15m/s 이상", "PM2.5 보정"],
             ["터널공사(옥내공사)", "-", "체감 33℃ 이상", "0℃ 이하", "-", "-", "PM2.5 보정"]],
            "비작업일수 산정을 위한 기상조건 적용 기준(가이드라인 예시)", size=8)
    D.table(["작업", "작업제한 기상조건", "근거"],
            [["콘크리트 타설", "강우·강설 시 원칙적 금지 / 일평균 4℃ 이하 한중콘크리트 / 일평균 25℃ 초과 서중콘크리트", "KCS 14 20 10·40·41"],
             ["타워크레인", "순간풍속 10m/s 이상 설치·해체 중지, 15m/s 이상 운전중지", "KCS 21 20 10, 산업안전보건기준 규칙 제37조"],
             ["굴착 및 복구공사 다짐", "0℃ 이하, 강우 시", "SMCS 11 85 10"],
             ["보강토 옹벽", "1.5℃ 미만, 강우·강설 시", "KCS 11 80 10"],
             ["아스팔트 포장(표층·기층)", "5℃ 이하, 강우 시 (프라임코트 10℃ 이하)", "KCS 44 50 10"],
             ["강구조 현장용접", "-20℃ 이하, 강우, 풍속 2m/s 이상", "KCS 14 31 30"],
             ["도장", "5℃ 미만, 강우·강풍·강설 시, 43℃ 이상", "KCS 14 31 40"],
             ["폭염작업", "체감 33℃ 이상 시 2시간마다 20분 이상 휴식 등", "산업안전보건기준 규칙 제560조"]],
            "주요 작업제한 기상조건(가이드라인 부록2 발췌)", widths=[4, 8.5, 4], align=["C", "L", "L"], size=8)
    pn = list(setup.profiles)
    D.table(["번호", "기상조건"] + pn, [[c, R.CONDITIONS[c]] + [setup.profiles[x].get(c) or "" for x in pn] for c in range(1, 15)],
            "금회 적용 기상프로필별 기상조건·적용률", size=7.5)
    D.table(["기상프로필", "설명", "적용 공정"],
            [[x, R.PROFILE_DESC.get(x, "사용자 정의"), ", ".join(a.id for a in fin.acts if a.profile == x) or "-"] for x in pn],
            "기상프로필 적용 공정", widths=[3, 7, 6.5], align=["C", "L", "L"], size=8)

    # 부록3
    D.new_section(SECTIONS[2], landscape=True)
    D.h2(SECTIONS[2])
    st = setup.station_data()
    rows = []
    for c in range(1, 15):
        v = [setup.cond_value(c, m) for m in range(1, 13)]
        rows.append([c, R.CONDITIONS[c]] + v + [round(sum(v), 1)])
    D.table(["번호", "기상조건"] + [f"{m}월" for m in range(1, 13)] + ["계"], rows,
            f"관측지점 '{setup.station}'({st['region']}, 지점 {st['code']}) 기상조건별 월평균 비작업일수 (2015~2024)", size=7.5)
    D.note("※ 1~13: 가이드라인 부록3 원자료, 14: 미세먼지 비상저감조치 발령일수(입력값)")
    mr = monthly_rows(fin, setup)
    for x in [x for x in pn if any(a.profile == x for a in fin.acts)]:
        rr = []
        for r in mr:
            mi = setup.month(x, r["년"], r["월"])
            rr.append([f"{r['년']}.{r['월']:02d}", mi.days, mi.A, mi.B, mi.C, mi.raw, mi.applied, mi.workable, f"{mi.rate * 100:.1f}%"])
        D.table(["년월", "달력일수", "기후(A)", "공휴일(B)", "중복(C)", "A+B−C", "적용(하한8)", "작업가능일", "가동률"], rr,
                f"기상프로필 '{x}' 월별 비작업일수 산정표", size=7.5)
    rr = [[f"{r['년']}.{r['월']:02d}", r["구분"], r["달력일수"], r["공사일수"], r["지배프로필"], r["적용"] if r["적용"] is not None else "-",
           r["반영비작업"], r["반영작업"]] for r in mr]
    rr.append(["계", "", sum(r["달력일수"] for r in mr), sum(r["공사일수"] for r in mr), "", "", round(sum(r["반영비작업"] for r in mr), 1),
               round(sum(r["반영작업"] for r in mr), 1)])
    D.table(["년월", "구분", "달력일수", "공사진행일", "주공정 기상프로필", "적용 비작업", "반영 비작업", "반영 작업일"], rr,
            "주공정 기준 공사기간 반영 비작업일수", size=7.5, bold_rows=(len(rr) - 1,))

    # 부록4
    D.new_section(SECTIONS[3], landscape=True)
    D.h2(SECTIONS[3])
    used = {}
    for v in res.final_items.values():
        for it in v:
            if it.db:
                used.setdefault(it.db["no"], (it.db, []))[1].append(it.no)
    D.table(["DB번호", "분야", "대공정", "세부공종", "작업조건/산출근거", "1일작업량", "출처", "적용 수량 No"],
            [[db["no"], db["field"], db["group"], db["name"], db["condition"], db["display"], db["source"], ", ".join(map(str, nos))]
             for db, nos in sorted(used.values(), key=lambda x: x[0]["no"])],
            "부록4(1일 작업량) DB 적용 항목", widths=[1.5, 1.8, 2.4, 5, 6, 3, 4, 3], align=["C", "C", "C", "L", "L", "C", "L", "L"], size=7.5)
    man = [it for v in res.final_items.values() for it in v if it.match_note == "수동입력"]
    if man:
        D.table(["No", "공정", "공종명", "규격", "단위", "1일작업량", "근거"],
                [[it.no, it.act, it.name, it.spec, it.unit, round(it.base_rate, 3), it.note or "품셈·실적"] for it in man],
                "품셈·실적 적용 항목(수동 입력)", widths=[1.2, 1.5, 6, 6, 1.5, 2.5, 8], align=["C", "C", "L", "L", "C", "R", "L"], size=7.5)
    auto = [it for v in res.final_items.values() for it in v if str(it.match_note).startswith(("규칙", "유사"))]
    if auto:
        D.table(["No", "공종명", "규격", "단위", "매칭 DB 공종", "1일작업량", "매칭방법"],
                [[it.no, it.name, it.spec, it.unit, it.db["name"], it.db["display"], it.match_note] for it in auto],
                "자동매칭 항목(검토자 확인)", widths=[1.2, 5.5, 5.5, 1.5, 5, 2.5, 5], align=["C", "L", "L", "C", "L", "C", "L"], size=7.5)

    # 부록5
    D.new_section(SECTIONS[4], landscape=True)
    D.h2(SECTIONS[4])
    D.para("작업일수 = ROUNDUP( 수량 ÷ (기준 1일작업량 × 보정계수 × 작업조) ),  보정계수 = 현장여건 × 공사규모 × 지질조건 × 기상기후 × 개별계수", indent=0)
    for a in fin.acts:
        its = res.final_items.get(a.id, [])
        if not its:
            D.para(f"[{a.id}] {a.field}·{a.name} — 직접입력 작업일수 {a.direct_days:g}일 ({a.note or '공정계획·실적 기준'})", indent=0, bold=True)
            continue
        rows = [[it.no, it.name, it.spec, it.unit, it.qty, it.base_rate, it.factor, it.crews, it.days, it.indirect or "", it.correction or "",
                 "중복" if it.parallel else it.total, it.formula] for it in its]
        rows.append(["", "공정 작업일수 계", "", "", "", "", "", "", "", "", "", a.work, f"직접입력 {a.direct_days:g}일 포함" if a.direct_days else ""])
        D.table(["No", "공종명", "규격", "단위", "수량", "기준1일작업량", "보정", "작업조", "순작업", "간접", "보정일", "합산", "산식"], rows,
                f"[{a.id}] {a.field} · {a.group} · {a.name}", widths=[1, 4.5, 4.2, 1.2, 1.8, 2, 1.2, 1.2, 1.3, 1.1, 1.1, 1.3, 5.2],
                align=["C", "L", "L", "C", "R", "R", "C", "C", "C", "C", "C", "C", "L"], size=7, bold_rows=(len(rows) - 1,))

    # 부록6
    D.new_section(SECTIONS[5], landscape=True)
    D.h2(SECTIONS[5])
    rows, fills = [], []
    for a in sorted(fin.acts, key=lambda a: (a.es, a.ef)):
        d = cal.day_to_date
        rows.append([a.id, a.field, a.name, a.profile, round(a.work), ", ".join(r.label() for r in a.preds) or "-",
                     fmt_date(d(a.es)), fmt_date(d(a.ef)), fmt_date(d(a.ls)), fmt_date(d(a.lf)), a.cal_days, a.tf, "●" if a.critical else ""])
        fills.append("FFE0E0" if a.critical else None)
    D.table(["ID", "분야", "공정명", "기상프로필", "작업일수", "선행관계", "ES", "EF", "LS", "LF", "공사기간", "TF", "CP"], rows,
            "CPM 계산표", widths=[1.2, 1.3, 5, 2, 1.5, 3.5, 2.1, 2.1, 2.1, 2.1, 1.6, 1.1, 1], size=7.5, fills=fills,
            align=["C", "C", "L", "C", "C", "C", "C", "C", "C", "C", "C", "C", "C"])
    D.note("※ ES/EF: 최조개시/최조완료, LS/LF: 최늦개시/최늦완료, TF(총여유)=LS−ES. 선행관계 FS=완료 후 착수, SS n%=선행 n% 진행 후 착수, FF=동시완료")
    D.note(f"※ 주공정: {' → '.join(fin.cp_ids)} / 준비기간 {fin.prep}일 이후 착수, 최종공정 완료 후 정리기간 {fin.cleanup}일")

    # 부록7
    D.new_section(SECTIONS[6])
    D.h2(SECTIONS[6])
    if res.target_days is None:
        D.para("목표기간 기준이 설정되지 않아 공기조정을 수행하지 않았음.", indent=0)
    else:
        D.table(["항목", "내용"], [["목표기간", f"{res.target_days}일 ({res.target_desc})"], ["기본 산정", f"{res.base.total_days}일"],
                                  ["최종", f"{fin.total_days}일"], ["목표 충족", "충족" if fin.total_days <= res.target_days else "미충족"]],
                "공기조정 개요", widths=[4, 12], align=["C", "L"])
        if res.compress_history:
            D.table(["단계", "구분", "조치내용", "조정 전", "조정 후", "단축"],
                    [[h["단계"], h["구분"], h["조치내용"], h["조정전(일)"], h["조정후(일)"], h["단축(일)"]] for h in res.compress_history],
                    "공기조정 단계별 이력", widths=[1.2, 1.5, 9.5, 1.6, 1.6, 1.3], align=["C", "C", "L", "C", "C", "C"], size=7.5)
        else:
            D.para("산정 공사기간이 목표기간 이내로 조정이 필요하지 않았음.", indent=0)

    # 부록8
    D.new_section(SECTIONS[7])
    D.h2(SECTIONS[7])
    adq = res.adequacy
    for r in adq["rows"]:
        D.h4(f"■ {r['구분']} — {r['기준']}")
        D.bullet(f"산정식 : {r['산정식']}")
        if r.get("변수"):
            D.bullet(f"변수 : {r['변수']}")
        if r.get("Y"):
            D.bullet(f"Y = {r['Y']:.1f}일")
        if r.get("B"):
            D.bullet(f"B = {r['B']:.0f}일 ({r['B'] / DPM:.1f}개월),  A = {adq['A']}일,  편차 = (A−B)/B = {r['편차'] * 100:+.1f}%  → {r['판정']}")
        if r.get("비고"):
            D.bullet(f"비고 : {r['비고']}")
    D.table(["시설물", "산정공식", "변수", "적용범위"], [[k, v[4], v[1], v[3]] for k, v in R.FORMULAS.items()],
            "가이드라인 부록5 시설물별 공사기간 산정공식", widths=[3, 6.5, 4.5, 2.5], align=["C", "L", "L", "C"], size=7.5)
    D.note("※ 산정값(Y)은 비작업일수를 포함하므로 준비기간·정리기간만 합산. 총공사비는 추정금액(추정가격+부가가치세+관급자재비). 적용범위 내에서만 활용")

    # 부록9
    D.new_section(SECTIONS[8])
    D.h2(SECTIONS[8])
    if res.winter:
        wi = res.winter
        D.table(["구분", "공사기간(일)", "개월", "착공", "준공"],
                [["동절기 미반영(금회 산정)", fin.total_days, round(fin.months, 1), fmt_date(cal.day_to_date(0)), fmt_date(fin.end_date)],
                 ["동절기 공사중지 반영", wi.total_days, round(wi.months, 1), fmt_date(wi.cal.day_to_date(0)), fmt_date(wi.end_date)],
                 ["증 감", wi.total_days - fin.total_days, round(wi.months - fin.months, 1), "", ""]],
                f"동절기({','.join(map(str, p.winter_months))}월) 공사중지 반영 시 공사일정", size=8.5)
        D.para("동절기 중지기간은 달력일수 전체를 비작업일수로 계상하여 기상조건에 의한 비작업일수를 대체함(중복계상 배제). "
               "가이드라인 부록5 산정공식(실적공기)과의 비교는 동절기 미반영 기준으로 수행하였으며, 실제 공사 시 동절기 공사중지 여부는 "
               "시공조건 명시사항에 기재하여야 함.", indent=0)
    else:
        D.para(f"동절기 공사중지 설정: '{p.winter_mode}'", indent=0)

    # 부록10
    D.new_section(SECTIONS[9])
    D.h2(SECTIONS[9])
    items6 = [("용지보상", ["미처리 용지의 장소, 범위 및 처리 예정 시기", "공사용지 등의 사용종료 후의 복구내용",
                        "공사용 가설도로·기자재 보관용 토지 임대 시 장소·범위·시기·기간·사용조건·복구방법"]),
              ("인·허가 및 심의 절차", ["관계기관·지자체 협의결과 특정조건이 부여되어 공정에 영향이 있는 경우 그 항목 및 영향범위",
                                   "협의 미완료 부분의 제약내용 및 완료 예정시기", "타 공사의 개시·완료 시기에 의한 영향",
                                   "시공시기·시간·방법이 제한되는 경우 그 내용"]),
              ("지장물 이설·철거", ["지하매설물 사전조사 항목 및 조사기간, 이설기간", "점용물건 지장물의 명칭·관리자·위치·이설시기·방호"]),
              ("문화재 시·발굴", ["사전조사 항목 및 조사기간", "시·발굴 예정기간"]),
              ("환경 관련", ["사후환경영향조사 항목 및 기간", "소음·진동·분진·배출가스 방지 시공방법·작업시간", "탁수·용수 처리대책",
                          "지반침하·지하수 고갈 등 사전·사후 조사", "매립폐기물·오염토 처리기간"]),
              ("안전대책", ["철도·가스·전기·통신·수도 근접공사의 시공방법·작업시간 제한", "낙석·붕괴 방호시설", "교통유도원·보안요원 배치",
                         "유독가스·산소결핍 대책 환기설비"]),
              ("공사용도로·가설·건설부산물·기타", ["반입로 사용 제한, 가설도로 안전시설·유지보수", "가설비 구조·시공방법 지정",
                                          "건설발생토 반출·가적치 조건, 폐기물 처리조건", "지급자재·대여품, 공사용 전력, 신기술·특허공법, 부분사용, 연약지반 개량"])]
    for k, lst in items6:
        D.h4(f"■ {k}")
        for t in lst:
            D.bullet(t)

    # 부록11
    D.new_section(SECTIONS[10], landscape=True)
    D.h2(SECTIONS[10])
    D.table(["시각", "단계", "구분", "내용"], [[e["시각"], e["단계"], e["구분"], e["내용"]] for e in res.log.entries],
            "검토 프로그램 실행 작업경과", widths=[1.6, 3.2, 1.3, 20.5], align=["C", "C", "C", "L"], size=7,
            fills=["FCE4D6" if e["구분"] == "경고" else None for e in res.log.entries])
    D.save(path)
