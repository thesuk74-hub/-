"""⑤ 주공정(Critical Path) 구성 — CPM 전진·후진 계산.

 · 공정별 작업일수(W)를 각 공정의 기상프로필 가동률 달력 위에서 소진하여 달력 공사기간 산정
 · 선후관계: FS(+lag일) / SS(+lag일 또는 선행 진행률 %) / FF(+lag일)
 · 총여유(TF) = 최늦개시 − 최조개시,  TF = 0 공정이 주공정
 · 공사기간 = 준비기간 + (주공정 착수~완료 달력일) + 정리기간
"""
import re
from dataclasses import dataclass, field

from . import refdata as R
from .calendar_cap import CapacityCalendar

REL_RE = re.compile(r"^\s*([^:\s]+)\s*(?::\s*(FS|SS|FF)\s*([+-]\s*\d+(?:\.\d+)?)?\s*(\d+(?:\.\d+)?\s*%)?)?\s*$", re.I)


@dataclass
class Rel:
    pred: str
    type: str = "FS"
    lag: float = 0.0        # 달력일
    pct: float = None       # SS 진행률(0~1)

    def label(self):
        if self.type == "SS" and self.pct is not None:
            return f"{self.pred}:SS{self.pct * 100:.0f}%"
        lag = f"{self.lag:+g}" if self.lag else ""
        return f"{self.pred}:{self.type}{lag}" if (self.type != "FS" or lag) else self.pred


def parse_preds(text):
    rels = []
    if not text:
        return rels
    for tok in re.split(r"[,;\n]", str(text)):
        tok = tok.strip()
        if not tok:
            continue
        m = REL_RE.match(tok.replace(" ", ""))
        if not m:
            raise ValueError(f"선행공정 표기 오류: '{tok}' (예: C10, C20:SS50%, C30:FS+5, C40:FF+0)")
        pid, typ, lag, pct = m.groups()
        r = Rel(pid, (typ or "FS").upper(), float(lag.replace(" ", "")) if lag else 0.0)
        if pct:
            r.pct = float(pct.rstrip("%")) / 100
            r.type = "SS"
        rels.append(r)
    return rels


@dataclass
class Activity:
    id: str
    field: str
    group: str
    name: str
    profile: str
    direct_days: float = 0
    preds: list = field(default_factory=list)
    max_overlap: float = 0.0          # 공기조정 시 선행공정과 중첩 가능 최대율(0~1)
    weight: float = None              # 공정률 가중치(공사비)
    milestone: str = ""
    note: str = ""
    snet: object = None               # 착수가능일(Start No Earlier Than) — 복합공사 후속 공사 착공일 등
    project: str = ""                 # 복합공사의 소속 공사명
    work: float = 0.0                 # 작업일수(W)
    # CPM 결과(일차, 준비기간 포함 착공일 기준 0일차)
    es: int = 0
    ef: int = 0
    ls: int = 0
    lf: int = 0
    tf: int = 0
    critical: bool = False

    @property
    def cal_days(self):
        return self.ef - self.es + 1 if self.work > 0 else 0


class CycleError(ValueError):
    pass


def topo_order(acts):
    ids = {a.id for a in acts}
    for a in acts:
        for r in a.preds:
            if r.pred not in ids:
                raise ValueError(f"공정 {a.id}의 선행공정 '{r.pred}'이(가) 공정표에 없습니다.")
    indeg = {a.id: len(a.preds) for a in acts}
    succ = {a.id: [] for a in acts}
    for a in acts:
        for r in a.preds:
            succ[r.pred].append(a.id)
    order, q = [], [a.id for a in acts if indeg[a.id] == 0]
    while q:
        x = q.pop(0)
        order.append(x)
        for s in succ[x]:
            indeg[s] -= 1
            if indeg[s] == 0:
                q.append(s)
    if len(order) != len(acts):
        raise CycleError("선행관계에 순환(cycle)이 있습니다: " + ", ".join(i for i in indeg if indeg[i] > 0))
    return order, succ


def cpm(acts, cal: CapacityCalendar, prep_days):
    A = {a.id: a for a in acts}
    order, succ = topo_order(acts)
    # 전진계산
    for i in order:
        a = A[i]
        es = prep_days
        if a.snet:
            es = max(es, cal.date_to_day(a.snet))
        ff_req = None
        for r in a.preds:
            p = A[r.pred]
            if r.type == "FS":
                es = max(es, p.ef + 1 + int(r.lag))
            elif r.type == "SS":
                if r.pct is not None:
                    pt = cal.finish(p.profile, p.es, p.work * r.pct) + 1 if p.work > 0 else p.es
                    es = max(es, pt + int(r.lag))
                else:
                    es = max(es, p.es + int(r.lag))
            elif r.type == "FF":
                req = p.ef + int(r.lag)
                ff_req = req if ff_req is None else max(ff_req, req)
        if ff_req is not None and a.work > 0:
            if cal.finish(a.profile, es, a.work) < ff_req:
                es = max(es, cal.start_for_finish(a.profile, ff_req, a.work))
        a.es = es
        a.ef = cal.finish(a.profile, es, a.work) if a.work > 0 else es - 1
    finish = max(a.ef for a in acts)
    # 후진계산
    for i in reversed(order):
        a = A[i]
        lf, ls_cap = finish, None
        for sid in succ[i]:
            s = A[sid]
            for r in s.preds:
                if r.pred != i:
                    continue
                if r.type == "FS":
                    lf = min(lf, s.ls - 1 - int(r.lag))
                elif r.type == "SS":
                    if r.pct is not None and a.work > 0:
                        cand = cal.start_for_finish(a.profile, s.ls - 1 - int(r.lag), a.work * r.pct)
                    else:
                        cand = s.ls - int(r.lag)
                    ls_cap = cand if ls_cap is None else min(ls_cap, cand)
                elif r.type == "FF":
                    lf = min(lf, s.lf - int(r.lag))
        if a.work > 0:
            ls = cal.start_for_finish(a.profile, lf, a.work)
            if ls_cap is not None and ls_cap < ls:
                ls = ls_cap
                lf = cal.finish(a.profile, ls, a.work)
        else:
            ls = lf + 1 if ls_cap is None else min(lf + 1, ls_cap)
        a.ls, a.lf = ls, lf
        a.tf = a.ls - a.es
        a.critical = a.tf <= 0
    return finish


@dataclass
class ScheduleResult:
    acts: list
    cal: CapacityCalendar
    prep: int
    cleanup: int
    finish_day: int            # 마지막 공정 종료일차
    total_days: int            # 준비 + 공사 + 정리 (달력일)
    cp_ids: list
    cp_work: float             # 주공정 순작업일수(달력분해)
    cp_nonwork: float          # 주공정 비작업일수(동절기 중지 제외)
    daily_profile: list        # 공사기간 일별 지배 프로필(주공정)
    prep_end: int = 0          # 준비기간 마지막 일차
    work_start: int = 0        # 공사 착수 일차
    winter_stop_days: float = 0  # 동절기 공사중지로 공정이 멈춘 일수(준비·정리 포함)
    winter_cal_days: int = 0     # 공사기간 내 동절기(중지월) 달력일수 — 엑셀 '동절기 제외' 비교용

    @property
    def months(self):
        from .refdata import DAYS_PER_MONTH
        return self.total_days / DAYS_PER_MONTH

    @property
    def end_date(self):
        return self.cal.day_to_date(self.total_days - 1)

    def act(self, i):
        return next(a for a in self.acts if a.id == i)


def critical_chain(acts):
    """주공정 공정 ID 목록(착수순)."""
    return [a.id for a in sorted(acts, key=lambda a: (a.es, a.ef)) if a.critical and a.work > 0]


def run_schedule(acts, cal, prep, cleanup):
    # 준비기간: 비작업일 미계상(달력일), 동절기 공사중지 시에는 중지월 제외
    prep_end = cal.finish(R.PREP_PROFILE, 0, prep) if prep > 0 else -1
    start = prep_end + 1
    finish = cpm(acts, cal, start)
    cp = critical_chain(acts)
    end = cal.finish(R.PREP_PROFILE, finish + 1, cleanup) if cleanup > 0 else finish
    total = end + 1
    setup = cal.setup
    wm = set(setup.winter_months) if setup.winter_stop else set()
    # 주공정 기준 일별 지배 프로필 → 비작업/작업일/동절기중지 분해
    daily, nonwork, wstop = [], 0.0, 0.0
    crit = [a for a in acts if a.critical and a.work > 0]
    for d in range(start, finish + 1):
        active = [a for a in crit if a.es <= d <= a.ef]
        if active:
            best = max(active, key=lambda a: cal.cap(a.profile, d))
            c = cal.cap(best.profile, d)
            daily.append(best.profile)
        else:
            c = 1.0 if not (wm and cal.dates[d].month in wm) else 0.0
            daily.append(None)
        if c == 0 and wm and cal.dates[d].month in wm:
            wstop += 1
        else:
            nonwork += 1 - c
    wstop += (start - prep) + (total - (finish + 1) - cleanup)
    span = finish + 1 - start
    wcal = sum(1 for d in range(total) if cal.dates[d].month in wm)
    return ScheduleResult(acts, cal, prep, cleanup, finish, total, cp, span - nonwork - (wstop - (start - prep) - (total - finish - 1 - cleanup)),
                          nonwork, daily, prep_end, start, wstop, wcal)
