"""검토 파이프라인 — 검토엑셀 ①→⑦ 순서로 실행하며 작업경과를 기록한다."""
import datetime as dt
from dataclasses import dataclass, field

from . import refdata as R
from .adequacy import evaluate
from .calendar_cap import CapacityCalendar
from .compress import compress
from .nonwork import NonworkSetup
from .project import read_project
from .schedule import run_schedule
from .workdays import activity_workdays


@dataclass
class RunLog:
    entries: list = field(default_factory=list)
    echo: bool = True

    def __call__(self, msg, step=None, level="정보"):
        e = {"시각": dt.datetime.now().strftime("%H:%M:%S"), "단계": step or (self.entries[-1]["단계"] if self.entries else ""),
             "구분": level, "내용": msg}
        self.entries.append(e)
        if self.echo:
            print(("[경고] " if level == "경고" else "") + msg, flush=True)

    def warn(self, msg):
        self(msg, level="경고")

    @property
    def warnings(self):
        return [e["내용"] for e in self.entries if e["구분"] == "경고"]


@dataclass
class Result:
    p: object
    log: RunLog
    setup: NonworkSetup
    cal: CapacityCalendar
    items_by_act: dict
    base: object                     # 기본 산정 ScheduleResult
    final: object                    # 최종 채택 ScheduleResult
    final_items: dict
    compress_history: list
    compress_ok: bool
    target_days: int
    target_desc: str
    adequacy: dict
    winter: object = None            # 동절기 반영 추가검토 ScheduleResult
    correction_days: int = 0

    @property
    def adopted_total(self):
        return self.final.total_days + self.correction_days


def build_items(p, setup, log):
    fac = p.factors()
    by_act = {a.id: [] for a in p.acts}
    A = {a.id: a for a in p.acts}
    for it in p.items:
        a = A[it.act]
        it.resolve(a.field, a.profile, fac)
        by_act[a.id].append(it)
        if it.warn:
            log.warn(f"수량 No.{it.no} '{it.name}': {it.warn}")
    for a in p.acts:
        a.work = activity_workdays(a, by_act[a.id])
        if a.work <= 0:
            log.warn(f"공정 {a.id} '{a.name}': 작업일수 0 — 수량 또는 직접작업일수 입력 필요(마일스톤으로 처리)")
    return by_act


def run(path, echo=True, log=None):
    log = log if log is not None else RunLog(echo=echo)
    S = ["① 공사개요", "② 비작업일수", "③ 1일작업량 매칭", "④ 작업일수", "⑤ 주공정(CPM)", "⑤-1 공기조정", "⑥ 공사기간 종합", "⑦ 적정성 검토", "추가검토(동절기)"]

    log(f"[1/9] {S[0]} — 입력서식 읽기: {path}", S[0])
    p = read_project(path)
    if str(p.g("준비·정리기간 동절기 중지", "Y")).upper().startswith("N"):
        p.winter_exempt = set(p.winter_exempt) | {R.PREP_PROFILE}
    for w in p.warnings:
        log.warn(w)
    missing = [k for k, ok in (("4_공정 (공정 1개 이상)", p.acts), ("총공사비(백만원)", p.total_cost > 0),
                               ("착공예정일", p.g("착공예정일"))) if not ok]
    if missing:
        raise SystemExit("입력 누락: " + ", ".join(missing) + " — 입력서식을 작성한 후 다시 실행하십시오.")
    log(f"  공사명: {p.name} / 총공사비 {p.total_cost:,.0f}백만원 / 착공 {p.start_date} / 준비 {p.prep_days}일·정리 {p.cleanup_days}일")
    log(f"  공정 {len(p.acts)}개, 수량항목 {len(p.items)}개, 유사사례 {len(p.cases)}건")

    log(f"[2/9] {S[1]} — 관측지점 '{p.g('기상관측지점')}' (부록3 2015~2024), 법정공휴일(부록1), 주40시간 하한", S[1])
    winter_stop = p.winter_mode == "반영"
    setup = NonworkSetup(station=p.g("기상관측지점", "서울"), profiles=p.profiles, dust=p.dust,
                         floor_days=float(p.g("주40시간 하한(일/월)", 8)),
                         round_int=str(p.g("비작업일수 정수반올림", "Y")).upper() == "Y",
                         winter_stop=winter_stop, winter_months=p.winter_months, no_holiday_profiles=p.holiday_off,
                         winter_exempt=p.winter_exempt, custom=p.custom, custom_label=p.custom_label)
    setup.station_data()
    for pn, conds in p.profiles.items():
        txt = ", ".join(f"{R.COND_SHORT[c]}×{w:g}" for c, w in sorted(conds.items())) or "기상조건 미적용"
        ann = setup.annual_climate(pn)
        tot = sum(sum(v) for v in ann.values())
        log(f"  프로필 '{pn}': {txt} → 연간 기상 비작업 {tot:.1f}일")
    cal = CapacityCalendar(setup, p.start_date)

    log(f"[3/9] {S[2]} — 수량 ↔ 부록4 1일작업량 (수동입력 > DB지정 > 자동매칭)", S[2])
    by_act = build_items(p, setup, log)
    n_auto = sum(1 for it in p.items if it.match_note.startswith(("규칙", "유사")))
    n_man = sum(1 for it in p.items if it.match_note == "수동입력")
    n_db = sum(1 for it in p.items if it.match_note == "DB지정")
    n_none = sum(1 for it in p.items if not it.rate)
    log(f"  수동입력 {n_man} / DB지정 {n_db} / 자동매칭 {n_auto} / 미결정 {n_none}")
    f = p.factors()
    log(f"  보정계수: 현장여건 '{f.site}'·공사규모 '{f.scale}'·지질 '{f.geology}'·기상기후 '{f.climate}'")

    log(f"[4/9] {S[3]} — 공정별 작업일수", S[3])
    for a in p.acts:
        log(f"  {a.id:<5} {a.field:<4} {a.name:<24} 작업일수 {a.work:>6.0f}일  ({a.profile})")

    log(f"[5/9] {S[4]} — CPM 전진·후진 계산 (공정별 가동률 달력)", S[4])
    base = run_schedule(p.acts, cal, p.prep_days, p.cleanup_days)
    log(f"  주공정: {' → '.join(base.cp_ids)}")
    log(f"  공사기간 {base.total_days}일 ({base.months:.1f}개월) = 준비 {p.prep_days} + 비작업 {base.cp_nonwork:.0f} + 작업 {base.cp_work:.0f} + 정리 {p.cleanup_days}"
        + (f" + 동절기 중지 {base.winter_stop_days:.0f}" if base.winter_stop_days else ""))
    if winter_stop:
        log(f"  동절기 공사중지 반영 — 공사기간 내 동절기 달력일수 {base.winter_cal_days}일, 동절기 제외 공사기간 {base.total_days - base.winter_cal_days}일")

    log(f"[6/9] {S[5]}", S[5])
    tdays, tdesc = p.target_days()
    final, final_items, hist, ok = base, by_act, [], True
    if tdays is None:
        log("  목표기간 기준 '없음' — 공기조정 생략")
    elif base.total_days <= tdays:
        log(f"  산정 {base.total_days}일 ≤ 목표 {tdays}일({tdesc}) — 조정 불필요")
    else:
        res, items2, hist, ok = compress(p.acts, by_act, cal, p.prep_days, p.cleanup_days, tdays, log)
        if str(p.g("공기조정 결과 채택", "Y")).upper() == "Y":
            final, final_items = res, items2
        else:
            log("  '공기조정 결과 채택'=N — 기본 산정안을 최종으로 유지(조정안은 참고)")
            final, final_items = base, by_act
            final.alt = res
        if not ok:
            log.warn(f"목표기간 {tdays}일 이내로 조정 불가 (최소 {res.total_days}일) — 목표기간 재검토 또는 공법·작업시간 조정 필요")

    log(f"[7/9] {S[6]}", S[6])
    corr = int(_num(p.g("공사여건 보정일수"), 0))
    total = final.total_days + corr
    log(f"  최종 공사기간 {total}일 ≒ {total / R.DAYS_PER_MONTH:.1f}개월 (착공 {p.start_date} ~ 준공 {final.cal.day_to_date(total - 1)})"
        + (f", 공사여건 보정 +{corr}일 포함" if corr else ""))

    log(f"[8/9] {S[7]} — 부록5 산정공식·동종공사 실적 ±{p.tolerance * 100:.0f}%", S[7])
    a_cmp = total - (final.winter_cal_days if winter_stop else 0)
    if winter_stop:
        log(f"  비교용 공사기간(A) = {total} − 동절기 {final.winter_cal_days} = {a_cmp}일 (부록5 실적공기는 동절기 공사중지 미포함)")
    adq = evaluate(p, a_cmp, p.prep_days, p.cleanup_days)
    adq["winter_excluded"] = final.winter_cal_days if winter_stop else 0
    for r in adq["rows"]:
        if r.get("B"):
            log(f"  {r['구분']:<10} B={r['B']:.0f}일({r['B'] / R.DAYS_PER_MONTH:.1f}개월) 편차 {r['편차'] * 100:+.1f}% → {r['판정']}")
        else:
            log(f"  {r['구분']:<10} {r['비고']}")
    log(f"  종합판정: {adq['verdict']} (주 비교기준: {adq['basis']})")

    winter = None
    if p.winter_mode == "추가검토":
        log(f"[9/9] {S[8]} — 동절기({','.join(map(str, p.winter_months))}월) 공사중지 반영 시 일정", S[8])
        s2 = NonworkSetup(station=setup.station, profiles=setup.profiles, dust=setup.dust, floor_days=setup.floor_days,
                          round_int=setup.round_int, winter_stop=True, winter_months=p.winter_months,
                          no_holiday_profiles=p.holiday_off, winter_exempt=p.winter_exempt, custom=p.custom,
                          custom_label=p.custom_label)
        cal2 = CapacityCalendar(s2, p.start_date)
        import copy
        acts2 = copy.deepcopy(final.acts)
        winter = run_schedule(acts2, cal2, p.prep_days, p.cleanup_days)
        log(f"  동절기 반영 공사기간 {winter.total_days}일 ({winter.months:.1f}개월), 준공 {winter.end_date} (+{winter.total_days - final.total_days}일), "
            f"동절기 달력일수 {winter.winter_cal_days}일 제외 시 {winter.total_days - winter.winter_cal_days}일")
    else:
        log(f"[9/9] {S[8]} — 동절기 공사중지 '{p.winter_mode}' 설정으로 생략", S[8])

    return Result(p, log, setup, cal, by_act, base, final, final_items, hist, ok, tdays, tdesc, adq, winter, corr)


def _num(v, d=0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return d
