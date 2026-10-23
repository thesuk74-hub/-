---
name: construction-period-review
description: 공사기간 적정성 검토 실행 — 「적정 공사기간 확보를 위한 가이드라인(2026)」 기준으로 입력서식을 작성·실행하여 공사기간 산정 보고서, 보고서 부록, 주공정(Critical Path) 공정표, 검토과정 워크북을 만든다. 사용자가 공사기간 적정성 검토, 적정공기 산정, 공사기간 산정 보고서, 주공정 공정표를 요청하거나 /construction-period-review 를 입력하면 사용.
---

# 공사기간 적정성 검토

이 스킬은 `.claude/agents/construction-period-reviewer.md`의 절차를 그대로 따른다. 먼저 그 파일을 읽고 0~5단계를 수행한다.

빠른 실행:

```bash
# 예제 실행(첨부 샘플보고서 사업)
python3 -m cpa run examples/예제_대죽공공폐수_입력.xlsx -o output/예제_대죽

# 신규 사업
cp templates/입력서식_템플릿.xlsx output/<사업명>/입력_<사업명>.xlsx   # 작성 후
python3 -m cpa run output/<사업명>/입력_<사업명>.xlsx -o output/<사업명>

# 공정별 수량파일만 있는 경우(수량첨부 검토)
python3 -m cpa qty --work 처리시설공사 토공.xlsx 구조물공.xlsx --work 관공사 관로공.xlsx --start 2027-03-01 --end 2028-10-31 --cost 9800 --loc "충청남도 서산시" -o output/<사업명>

# 보조 명령
python3 -m cpa stations 충청남도     # 기상관측지점
python3 -m cpa db 콘크리트 타설      # 부록4 1일작업량 DB
python3 -m cpa fill 입력.xlsx --items 수량.csv --acts 공정.csv --cases 사례.csv
```

인자가 주어지면(예: 입력파일 경로) 그 파일로 바로 2단계(실행)부터 진행하고, 3단계 검토와 5단계 보고를 반드시 수행한다.
