@echo off
chcp 65001 > nul
cd /d %~dp0
echo [공사기간 적정성 검토] 필요한 패키지를 확인합니다...
python -m pip install -q -r requirements.txt
echo 브라우저에서 http://localhost:8000 을 엽니다. (종료: 이 창에서 Ctrl+C)
start "" http://localhost:8000
python -m cpa web --port 8000
pause
