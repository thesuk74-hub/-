"""검토 작업 단계 함수 — CLI(__main__)와 웹(web/app.py)이 공통으로 사용.

calc(입력) → chart(결과) → reports(결과) → export_pdf(폴더) / summary(결과)
"""
import json
import os
import shutil
import subprocess
import zipfile

from . import refdata as R
from .pipeline import run
from .summary import fmt_date, monthly_rows, progress_curve

def soffice():
    """LibreOffice 실행파일 (PATH 또는 Windows/macOS 기본 설치경로)."""
    for c in (shutil.which("soffice"), shutil.which("libreoffice"),
              r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if c and os.path.exists(c):
            return c
    return None


REPORT_FILES = ["01_공사기간산정보고서.docx", "02_보고서부록.docx", "03_주공정공정표.xlsx", "04_검토과정_워크북.xlsx"]


def calc(path, log=None):
    """①~⑦ 검토 실행. log: pipeline.RunLog(웹에서 진행상황 수집용) 또는 None."""
    return run(path, echo=log is None, log=log)


def calc_composite(combined_path, subs, log):
    """복합공사: 통합 입력 산정 + 공사별 입력 독립 산정(적정성). subs: [{"name","path"}]."""
    from .pipeline import RunLog
    log(f"[복합공사] 공사 {len(subs)}개 — 통합 공정망 산정", "복합공사")
    res = calc(combined_path, log)
    res.subs = []
    for k, s in enumerate(subs, 1):
        log(f"[복합공사 {k}/{len(subs)}] '{s['name']}' 공사별 독립 검토(적정성)", "공사별 검토")
        sl = RunLog(echo=False)
        r = calc(s["path"], sl)
        for w in sl.warnings:
            log.warn(f"[{s['name']}] {w}")
        mine = [a for a in res.final.acts if a.project == s["name"]]
        span = (min(a.es for a in mine if a.work > 0), max(a.ef for a in mine if a.work > 0)) if mine else None
        res.subs.append(dict(name=s["name"], path=s["path"], res=r, span=span))
        log(f"  공사기간 {r.adopted_total:,}일, 적정성 '{r.adequacy['verdict']}'", "공사별 검토")
    vs = [s["res"].adequacy["verdict"] for s in res.subs]
    ok = sum(1 for v in vs if v.startswith("적정"))
    ng = sum(1 for v in vs if v.startswith("재검토"))
    res.adequacy["verdict"] = ("적정" if ok == len(vs) else f"재검토 필요({ng}개 공사)" if ng else "일부 판정불가") + \
        f" — 공사별: 적정 {ok} / 재검토 {ng} / 기타 {len(vs) - ok - ng}"
    res.adequacy["basis"] = "공사별 독립 검토(부록5·동종실적)"
    log(f"[복합공사] 공사별 적정성 종합: {res.adequacy['verdict']}", "복합공사")
    return res


def sub_rows(res):
    """복합공사 공사별 요약표 행."""
    out = []
    cal = res.final.cal
    for s in getattr(res, "subs", []) or []:
        r = s["res"]
        prim = [x for x in r.adequacy["rows"] if x.get("주기준") and x.get("B")] or [x for x in r.adequacy["rows"] if x.get("B")]
        b = prim[0] if prim else None
        out.append({"공사명": s["name"], "통합_착수": str(cal.day_to_date(s["span"][0])) if s["span"] else "-",
                    "통합_완료": str(cal.day_to_date(s["span"][1])) if s["span"] else "-",
                    "공사비_백만원": r.p.total_cost, "시설유형": r.p.facility_type or "-",
                    "독립_공사기간_일": r.adopted_total, "A_일": r.adequacy["A"],
                    "B_기준": f"{b['구분']}({b['기준']})" if b else "-", "B_일": round(b["B"]) if b else None,
                    "편차_%": round(b["편차"] * 100, 1) if b else None, "판정": r.adequacy["verdict"],
                    "공정수": len(r.final.acts), "수량항목": sum(len(v) for v in r.final_items.values())})
    return out


def chart(res, out):
    from . import gantt, network
    os.makedirs(out, exist_ok=True)
    png, pdf = os.path.join(out, "03_주공정공정표.png"), os.path.join(out, "03_주공정공정표.pdf")
    gantt.render(res.final, res.p, png, pdf)
    network.render(res.final, os.path.join(out, "_주공정모식도.png"))
    return png


def reports(res, out):
    import openpyxl

    from . import appendix, gantt, report, workbook
    log = res.log
    os.makedirs(out, exist_ok=True)
    png = os.path.join(out, "03_주공정공정표.png")
    if not os.path.exists(png):
        chart(res, out)
    wb = openpyxl.Workbook()
    wb.active.title = "주공정공정표"
    gantt.write_sheet(wb.active, res.final, res.p)
    wb.save(os.path.join(out, "03_주공정공정표.xlsx"))
    log("  03_주공정공정표 (xlsx·pdf·png) 작성", "보고서")
    report.build(res, os.path.join(out, "01_공사기간산정보고서.docx"), png, os.path.join(out, "_주공정모식도.png"))
    log("  01_공사기간산정보고서.docx 작성", "보고서")
    appendix.build(res, os.path.join(out, "02_보고서부록.docx"))
    log("  02_보고서부록.docx 작성", "보고서")
    workbook.write(res, os.path.join(out, "04_검토과정_워크북.xlsx"))
    log("  04_검토과정_워크북.xlsx 작성", "보고서")
    for s_ in getattr(res, "subs", []) or []:
        so = os.path.join(out, "공사별", s_["name"].replace("/", "_").replace("·", ""))
        os.makedirs(so, exist_ok=True)
        r = s_["res"]
        log(f"  [공사별] {s_['name']} 보고서 작성", "보고서")
        chart(r, so)
        report.build(r, os.path.join(so, "01_공사기간산정보고서.docx"), os.path.join(so, "03_주공정공정표.png"),
                     os.path.join(so, "_주공정모식도.png"))
        appendix.build(r, os.path.join(so, "02_보고서부록.docx"))
        workbook.write(r, os.path.join(so, "04_검토과정_워크북.xlsx"))
    s = summary(res)
    with open(os.path.join(out, "result.json"), "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=1, default=str)
    with open(os.path.join(out, "작업경과.log"), "w", encoding="utf-8") as f:
        for e in log.entries:
            f.write(f"{e['시각']} [{e['단계']}] {'[경고] ' if e['구분'] == '경고' else ''}{e['내용']}\n")
    return s


def export_pdf(out, log=None):
    """보고서·부록 DOCX → PDF (LibreOffice) 및 전체 ZIP."""
    exe = soffice()
    made = []
    targets = [os.path.join(out, f) for f in ("01_공사기간산정보고서.docx", "02_보고서부록.docx")]
    sub = os.path.join(out, "공사별")
    if os.path.isdir(sub):
        for d in sorted(os.listdir(sub)):
            targets += [os.path.join(sub, d, f) for f in ("01_공사기간산정보고서.docx", "02_보고서부록.docx")]
    for src in targets:
        out_d = os.path.dirname(src)
        if not os.path.exists(src):
            continue
        if exe:
            subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", out_d, src],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
            pdf = src[:-5] + ".pdf"
            if os.path.exists(pdf):
                made.append(os.path.relpath(pdf, out))
                log and log(f"  PDF 변환: {os.path.relpath(pdf, out)}", "출력")
        else:
            log and log("  LibreOffice가 없어 PDF 변환 생략(DOCX 사용)", "출력")
    zpath = os.path.join(out, "성과품_전체.zip")
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for dp, _, fs in os.walk(out):
            for f in sorted(fs):
                if f.startswith("_") or f.endswith(".zip"):
                    continue
                full = os.path.join(dp, f)
                z.write(full, os.path.relpath(full, out))
    log and log("  성과품_전체.zip 작성", "출력")
    return made


def summary(res):
    """결과 확인 화면·result.json용 요약."""
    p, fin = res.p, res.final
    cal = fin.cal
    total = res.adopted_total
    adq = res.adequacy
    acts = []
    for a in sorted(fin.acts, key=lambda a: (a.es, a.ef)):
        acts.append(dict(id=a.id, 분야=a.field, 대공정=a.group, 공정명=a.name, 프로필=a.profile, 작업일수=round(a.work, 1),
                         선행=", ".join(r.label() for r in a.preds), 착수=str(cal.day_to_date(a.es)), 완료=str(cal.day_to_date(a.ef)),
                         기간=a.cal_days, TF=a.tf, CP=a.critical))
    items = []
    for v in res.final_items.values():
        for it in v:
            if it.warn or str(it.match_note).startswith(("규칙", "유사")):
                items.append(dict(No=it.no, 공정=it.act, 공종=it.name, 규격=it.spec, 단위=it.unit, 결정=str(it.match_note).split("(")[0] + (f"({it.confidence:.2f})" if str(it.match_note).startswith("유사") else ""),
                                  DB=(it.db or {}).get("name"), 일작업량=(it.db or {}).get("display") or it.base_rate, 경고=it.warn))
    return {
        "공사명": p.name, "발주청": p.g("발주청", ""), "착공": str(cal.day_to_date(0)), "준공": str(cal.day_to_date(total - 1)),
        "공사기간_일": total, "공사기간_개월": round(total / R.DAYS_PER_MONTH, 1),
        "준비기간": fin.prep, "비작업일수": round(fin.cp_nonwork), "작업일수": round(fin.cp_work), "정리기간": fin.cleanup,
        "동절기중지": round(fin.winter_stop_days), "공사여건보정": res.correction_days,
        "동절기제외_공사기간": total - fin.winter_cal_days if fin.winter_cal_days else None,
        "관측지점": res.setup.station, "주공정": fin.cp_ids, "주공정명": [fin.act(i).name for i in fin.cp_ids],
        "목표기간_일": res.target_days, "목표기간_설명": res.target_desc, "목표충족": (None if not res.target_days else total <= res.target_days),
        "기본산정_일": res.base.total_days, "공기조정": res.compress_history,
        "적정성": adq["verdict"], "적정성_기준": adq["basis"], "적정성_A": adq["A"], "허용범위": adq["tol"],
        "비교기준": [{"구분": r["구분"], "기준": r["기준"], "산정식": r.get("산정식"), "B_일": round(r["B"]) if r.get("B") else None,
                    "B_개월": round(r["B"] / R.DAYS_PER_MONTH, 1) if r.get("B") else None,
                    "편차_%": round(r["편차"] * 100, 1) if r.get("편차") is not None else None,
                    "판정": r.get("판정"), "주기준": r.get("주기준"), "비고": r.get("비고")} for r in adq["rows"]],
        "동절기반영_일": res.winter.total_days if res.winter else None,
        "동절기반영_준공": str(res.winter.end_date) if res.winter else None,
        "월별비작업": [{k: (round(v, 3) if isinstance(v, float) else v) for k, v in r.items()} for r in monthly_rows(fin, res.setup)],
        "공정률": [dict(년=y, 월=m, 월간=round(a, 1), 누계=round(b, 1)) for y, m, a, b in progress_curve(fin)],
        "공정": acts, "확인필요_수량": items, "경고": res.log.warnings,
        "수량항목수": sum(len(v) for v in res.final_items.values()),
        "복합공사": sub_rows(res),
    }
