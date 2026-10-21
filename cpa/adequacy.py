"""⑦ 공사기간 적정성 검토 — 가이드라인 제3장 2.(3)

  공사기간(A) : 준비기간 + 비작업일수 + 작업일수 + 정리기간 (본 산정)
  공사기간(B) : 부록5 시설물별 산정공식 Y 또는 최근 5년 동종공사 실적 평균 + 준비기간 + 정리기간
  |A − B| / B ≤ 20%  → 적정,  초과 시 산정과정 재검토
참고기준: 건설사업관리기간(공사비 기준), 한국환경공단 기준(환경시설)
"""
from . import refdata as R

DPM = R.DAYS_PER_MONTH


def _judge(A, B, tol):
    dev = (A - B) / B
    return dev, ("적정" if abs(dev) <= tol + 1e-9 else "재검토 필요")


def evaluate(p, A_days, prep, cleanup):
    tol = p.tolerance
    cost_m = p.total_cost            # 백만원
    cost_e = cost_m / 100            # 억원
    v = {"C억": cost_e, "C백만": cost_m, **p.formula_vars}
    rows = []

    # 1) 부록5 산정공식
    ftype = p.facility_type
    if ftype and ftype in R.FORMULAS:
        fn, vdesc, rng, rng_txt, ftxt = R.FORMULAS[ftype]
        need = R.FORMULA_VARS.get(ftype, [])
        miss = [k for k in need if not v.get(k)]
        if miss:
            rows.append(dict(구분="부록5 산정공식", 기준=f"{ftype}", 산정식=ftxt, 변수=vdesc, Y=None, B=None,
                             비고=f"변수 미입력({', '.join(miss)}) — 산정 불가", 주기준=False))
        else:
            Y = fn(v)
            B = Y + prep + cleanup
            inrange = rng(cost_e)
            dev, j = _judge(A_days, B, tol)
            rows.append(dict(구분="부록5 산정공식", 기준=f"{ftype}", 산정식=ftxt, 변수=vdesc, Y=Y, B=B, 편차=dev,
                             판정=j if inrange else f"{j}(참고)", 주기준=inrange,
                             비고=("" if inrange else f"총공사비 {cost_e:,.1f}억원 — 적용범위({rng_txt}) 외: 참고값")))
    # 2) 최근 5년 동종(유사)공사 실적 평균
    if p.cases:
        months = [c["공사기간(개월)"] for c in p.cases if c.get("공사기간(개월)")]
        if months:
            avg = sum(months) / len(months)
            Y = avg * DPM
            add = (prep + cleanup) if p.cases_add_prep else 0
            B = Y + add
            dev, j = _judge(A_days, B, tol)
            rows.append(dict(구분="동종공사 실적", 기준=f"최근 준공 동종·유사공사 {len(months)}건 평균 {avg:.1f}개월",
                             산정식="실적 공사기간 평균" + (" + 준비·정리기간" if add else ""), 변수="", Y=Y, B=B,
                             편차=dev, 판정=j, 주기준=True, 비고=""))
    # 3) 건설사업관리기간 (참고)
    cm = R.cm_period_months(cost_e)
    B = cm * DPM
    dev, j = _judge(A_days, B, tol)
    rows.append(dict(구분="건설사업관리기간", 기준=R.CM_SOURCE, 산정식=f"공사비 {cost_e:,.1f}억원 직선보간 → {cm:.1f}개월",
                     변수="", Y=None, B=B, 편차=dev, 판정=f"{j}(참고)", 주기준=False, 비고="참고"))
    # 4) 한국환경공단 산정식 (환경시설)
    if p.keco_type and p.keco_type in R.KECO_FORMULAS:
        fn, ftxt, vdesc, src = R.KECO_FORMULAS[p.keco_type]
        if v.get("V"):
            Y = fn(v)
            B = Y + prep + cleanup
            dev, j = _judge(A_days, B, tol)
            rows.append(dict(구분="환경공단 산정식", 기준=f"{p.keco_type} — {src}", 산정식=ftxt, 변수=vdesc, Y=Y, B=B,
                             편차=dev, 판정=f"{j}(참고)", 주기준=False, 비고="참고"))
    if p.keco_exp and p.keco_exp in R.KECO_EXPERIENCE:
        tab = R.KECO_EXPERIENCE[p.keco_exp]
        m = next((mm for lim, mm in tab if cost_e <= lim), tab[-1][1])
        B = m * DPM
        dev, j = _judge(A_days, B, tol)
        rows.append(dict(구분="환경공단 경험치", 기준=f"{p.keco_exp} — {R.KECO_EXP_SOURCE}", 산정식=f"공사비 {cost_e:,.1f}억원 → {m}개월",
                         변수="", Y=None, B=B, 편차=dev, 판정=f"{j}(참고)", 주기준=False, 비고="참고"))

    prim = [r for r in rows if r.get("주기준") and r.get("B")]
    mode = p.g("적정성 판정기준", "자동")
    if mode == "부록5":
        prim = [r for r in prim if r["구분"] == "부록5 산정공식"]
    elif mode == "동종공사":
        prim = [r for r in prim if r["구분"] == "동종공사 실적"]
    if prim:
        oks = [abs(r["편차"]) <= tol + 1e-9 for r in prim]
        ok = all(oks) if mode == "모두" else any(oks)
        verdict = "적정" if ok else "재검토 필요"
        basis = ", ".join(r["구분"] for r in prim)
        if ok and not all(oks):
            verdict = "적정(일부 기준 편차 초과 — 사유 검토)"
    else:
        verdict, basis = "판정불가(비교기준 없음)", "-"
    return dict(rows=rows, verdict=verdict, basis=basis, tol=tol, A=A_days)
