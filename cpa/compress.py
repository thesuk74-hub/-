"""⑤-1 공사기간 조정 — 산정 공사기간이 목표기간(발주처 요구기간 또는 건설사업관리기간)을 초과할 때.

탐욕적(greedy) 반복:
  1) 주공정 공정의 단위작업 중 작업일수가 가장 큰 항목에 작업조 +1 (최대작업조 이내)
  2) 동시수행 가능 공정(최대중첩률 > 0)의 주공정 선행관계를 FS → SS(선행 진행률%)로 10%p씩 중첩
매 단계 후보를 모두 시험 계산하여 총 공사기간 단축이 가장 큰 조치를 채택하고 경과를 기록한다.
"""
import copy

from .schedule import Rel, run_schedule
from .workdays import activity_workdays

STEP = 0.10


def _apply_work(acts, items_by_act):
    for a in acts:
        a.work = activity_workdays(a, items_by_act.get(a.id, []))


def _candidates(res, items_by_act):
    A = {a.id: a for a in res.acts}
    cands = []
    for a in res.acts:
        if not a.critical or a.work <= 0:
            continue
        its = [it for it in items_by_act.get(a.id, []) if not it.parallel and it.rate and it.max_crews
               and it.crews < it.max_crews and it.days > 1]
        if its:
            it = max(its, key=lambda x: x.days)
            cands.append(("작업조", a.id, it.no, f"[{a.id}] {a.name} — '{it.name}' 작업조 {it.crews}→{it.crews + 1}조"))
        if a.max_overlap > 0:
            for k, r in enumerate(a.preds):
                p = A[r.pred]
                if not p.critical or r.type == "FF":
                    continue
                cur = 0.0 if r.type == "FS" else (1 - r.pct if r.pct is not None else None)
                if cur is None:
                    continue
                new = round(cur + STEP, 2)
                if new <= a.max_overlap + 1e-9:
                    cands.append(("중첩", a.id, k, f"[{a.id}] {a.name} ← [{p.id}] {p.name} 중첩 {cur * 100:.0f}%→{new * 100:.0f}% (SS{(1 - new) * 100:.0f}%)"))
    return cands


def _apply(cand, acts, items_by_act):
    kind, aid, key, _ = cand
    if kind == "작업조":
        it = next(i for i in items_by_act[aid] if i.no == key)
        it.crews += 1
        it.recalc()
    else:
        a = next(x for x in acts if x.id == aid)
        r = a.preds[key]
        cur = 0.0 if r.type == "FS" else 1 - r.pct
        a.preds[key] = Rel(r.pred, "SS", 0.0, round(1 - (cur + STEP), 2))


def compress(acts, items_by_act, cal, prep, cleanup, target_days, log, max_iter=300):
    acts = copy.deepcopy(acts)
    items_by_act = copy.deepcopy(items_by_act)
    _apply_work(acts, items_by_act)
    res = run_schedule(acts, cal, prep, cleanup)
    history = []
    best = (res.total_days, copy.deepcopy(acts), copy.deepcopy(items_by_act))
    best_len, stall = 0, 0
    log(f"공기조정 시작: 산정 {res.total_days}일 → 목표 {target_days}일 (초과 {res.total_days - target_days}일)")
    for step in range(1, max_iter + 1):
        if res.total_days <= target_days:
            break
        cands = _candidates(res, items_by_act)
        if not cands:
            log("  더 이상 적용 가능한 작업조 증대·공정 중첩 조치가 없음")
            break
        trials = []
        for c in cands:
            a2, i2 = copy.deepcopy(acts), copy.deepcopy(items_by_act)
            _apply(c, a2, i2)
            _apply_work(a2, i2)
            r2 = run_schedule(a2, cal, prep, cleanup)
            gain = res.total_days - r2.total_days
            wsum = sum(x.work for x in r2.acts if x.critical)
            trials.append((gain, 1 if c[0] == "작업조" else 0, -wsum, c, a2, i2, r2))
        trials.sort(key=lambda t: (t[0], t[1], t[2]), reverse=True)
        gain, _, _, c, acts, items_by_act, r2 = trials[0]
        history.append({"단계": step, "구분": c[0], "조치내용": c[3], "조정전(일)": res.total_days,
                        "조정후(일)": r2.total_days, "단축(일)": gain})
        log(f"  {step:>3}. {c[3]} : {res.total_days}일 → {r2.total_days}일")
        res = r2
        if res.total_days < best[0]:
            best = (res.total_days, copy.deepcopy(acts), copy.deepcopy(items_by_act))
            best_len, stall = len(history), 0
        else:
            stall += 1
            if stall >= 6:
                log("  연속 6회 단축효과 없음 → 조정 종료")
                break
    total, acts, items_by_act = best
    if len(history) > best_len:
        log(f"  단축효과 없는 후속 {len(history) - best_len}개 조치는 미채택(최적 상태로 복원)")
        history = history[:best_len]
    _apply_work(acts, items_by_act)
    res = run_schedule(acts, cal, prep, cleanup)
    ok = res.total_days <= target_days
    log(f"공기조정 결과: {res.total_days}일 ({'목표 충족' if ok else '목표 미충족'})")
    return res, items_by_act, history, ok
