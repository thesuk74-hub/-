# 공사기간 적정성 검토 에이전트 — 프로젝트 지침

- 공사기간 적정성 검토·공사기간 산정 보고서·주공정 공정표 요청은 `.claude/agents/construction-period-reviewer.md` 절차로 수행한다(스킬 `/construction-period-review`).
- 계산은 `python3 -m cpa run <입력.xlsx> -o output/<사업명>` 으로만 수행하고, 결과 수치를 임의 수정하지 않는다. 입력을 고쳐 재실행한다.
- 기준: 「적정 공사기간 확보를 위한 가이드라인(2026, 국토교통부)」 원문 `cpa/data/reference_guideline_2026.pdf`.
- 수량첨부 검토(웹 기본 화면·`python3 -m cpa qty`): `qtyparse.py`(수량파일 해석) → `wizard.py`(공정 자동구성·간접/보정일수 → 입력서식) → 아래 엔진. 웹: `web/app.py`(공통 API)·`web/wiz.py`(수량첨부 API)·`web/static/index.html`(기본)·`classic.html`(입력서식 업로드).
- 성과품: 보고서 HWPX(`hwpxbuild.py`)·부록 XLSX(`xlsxbuild.py`) — 본문은 `report.py`·`appendix.py`를 `doc_cls`로 공유, PDF는 숨김 DOCX(`docbuild.py`)를 변환.
- 코드 구조: `refdata.py`(기준값) → `project.py`(입력서식) → `nonwork.py`·`calendar_cap.py`(②) → `workdays.py`(③④) → `schedule.py`(⑤ CPM) → `compress.py`(⑤-1) → `adequacy.py`(⑦) → `pipeline.py`(순서·작업경과) → `report.py`·`appendix.py`·`gantt.py`·`network.py`·`workbook.py`(성과품).
- 변경 후 확인: `python3 tools/make_example.py && python3 -m cpa run examples/예제_대죽공공폐수_입력.xlsx -o output/예제_대죽`
