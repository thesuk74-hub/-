"""출력물 공통 집계: 월별 비작업일수 표(표 2-6 형식), 공정률 곡선, 공종 집계."""
import calendar
import datetime as dt
from collections import OrderedDict, defaultdict

from . import refdata as R


def month_iter(d0, d1):
    y, m = d0.year, d0.month
    while (y, m) <= (d1.year, d1.month):
        yield y, m
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)


def monthly_rows(res, setup):
    """공사기간 전체 월별 비작업일수 산정표."""
    cal = res.cal
    start, end = cal.day_to_date(0), cal.day_to_date(res.total_days - 1)
    finish = res.finish_day
    rows = []
    for y, m in month_iter(start, end):
        nd = calendar.monthrange(y, m)[1]
        d0 = max(cal.date_to_day(dt.date(y, m, 1)), 0)
        d1 = min(cal.date_to_day(dt.date(y, m, nd)), res.total_days - 1)
        prep = sum(1 for d in range(d0, d1 + 1) if d < res.work_start)
        clean = sum(1 for d in range(d0, d1 + 1) if d > finish)
        work_idx = [d for d in range(d0, d1 + 1) if res.work_start <= d <= finish]
        profs = defaultdict(int)
        nonw = 0.0
        for d in work_idx:
            pf = res.daily_profile[d - res.work_start]
            profs[pf] += 1
            nonw += 1 - (cal.cap(pf, d) if pf else 1.0)
        gov = max(profs, key=profs.get) if profs else None
        info = setup.month(gov, y, m) if gov else None
        label = []
        if prep:
            label.append("준비")
        if work_idx:
            label.append("시운전" if gov == "연속작업" else "공사")
        if clean:
            label.append("정리")
        rows.append(OrderedDict(
            년=y, 월=m, 구분="·".join(label), 달력일수=nd, 대상일수=d1 - d0 + 1, 준비일수=prep, 정리일수=clean,
            공사일수=len(work_idx), 지배프로필=gov or "-",
            A=info.A if info else None, B=info.B if info else None, C=info.C if info else None,
            소계=info.raw if info else None, 적용=info.applied if info else None,
            작업가능=info.workable if info else None, 가동률=info.rate if info else None,
            반영비작업=round(nonw, 1), 반영작업=round(len(work_idx) - nonw, 1)))
    return rows


def progress_curve(res):
    """월별 계획 공정률(가중치=공사비, 미입력 시 작업일수). 공정기간 동안 균등 배분."""
    cal = res.cal
    acts = [a for a in res.acts if a.work > 0]
    w = {a.id: (a.weight if a.weight else a.work) for a in acts}
    tot = sum(w.values()) or 1
    start, end = cal.day_to_date(0), cal.day_to_date(res.total_days - 1)
    out, cum = [], 0.0
    for y, m in month_iter(start, end):
        nd = calendar.monthrange(y, m)[1]
        d0, d1 = cal.date_to_day(dt.date(y, m, 1)), cal.date_to_day(dt.date(y, m, nd))
        s = 0.0
        for a in acts:
            ov = max(0, min(a.ef, d1) - max(a.es, d0) + 1)
            if ov:
                s += w[a.id] * ov / (a.ef - a.es + 1)
        pct = s / tot * 100
        cum += pct
        out.append((y, m, pct, min(cum, 100.0)))
    return out


def group_summary(res, items_by_act):
    """분야·대공정별 작업일수 집계 (표 2-8 형식)."""
    g = OrderedDict()
    for a in res.acts:
        key = (a.field, a.group)
        d = g.setdefault(key, {"분야": a.field, "대공정": a.group, "공정": [], "작업일수": 0, "주공정": False})
        d["공정"].append(a.name)
        d["작업일수"] += a.work
        d["주공정"] |= a.critical
    return list(g.values())


def fmt_date(d):
    return d.strftime("%Y.%m.%d")


def months(days):
    return days / R.DAYS_PER_MONTH
