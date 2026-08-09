@echo off
chcp 65001 >nul
title DART MCP 서버 설치
cd /d "%~dp0"

echo.
echo  ================================================
echo    DART 조회 도구 설치 (Claude 데스크톱 앱용)
echo  ================================================
echo.
echo   설치가 끝나면 Claude 데스크톱 앱에서 이렇게 물어볼 수 있습니다.
echo.
echo     "게임회사들 매출채권회전율 비교해줘"
echo     "삼성전자 최근 5년 타법인출자 정리해줘"
echo     "이 회사 감사보고서에서 매출채권 주석 찾아줘"
echo.

REM ---------- 1. 파이썬 찾기 ----------
set "PY="
py -3 --version >nul 2>&1
if %errorlevel%==0 set "PY=py -3"
if defined PY goto HAVE_PY

python --version >nul 2>&1
if %errorlevel%==0 set "PY=python"
if defined PY goto HAVE_PY

echo  [오류] 파이썬이 설치되어 있지 않습니다.
echo.
echo   1) https://www.python.org/downloads/  접속
echo   2) 노란색 "Download Python" 버튼 클릭
echo   3) 설치 화면 맨 아래 "Add python.exe to PATH" 체크  ^<-- 중요!
echo   4) 설치가 끝나면 이 파일을 다시 더블클릭
echo.
pause
exit /b 1

:HAVE_PY
echo  [1/4] 파이썬 확인 완료
echo.

REM ---------- 2. API 키 ----------
echo  [2/4] OpenDART API 키가 필요합니다.
echo.
echo        아직 없으시면 https://opendart.fss.or.kr 에서
echo        무료로 발급받으실 수 있습니다 (이메일 인증 후 바로 발급).
echo.
set "DARTKEY="
set /p DARTKEY="        발급받은 API 키를 붙여넣고 Enter: "
if not defined DARTKEY goto NO_KEY
echo.

REM ---------- 3. 실행 환경 준비 ----------
REM 앱(.venv)과 분리한다. mcp 패키지가 streamlit과 버전이 충돌하기 때문이다.
if exist ".venv-mcp\Scripts\python.exe" goto HAVE_VENV

echo  [3/4] 최초 실행입니다. 준비하는 데 2~5분 걸립니다.
echo        (다음부터는 이 과정을 건너뜁니다)
echo.
%PY% -m venv .venv-mcp
if errorlevel 1 goto VENV_FAIL

:HAVE_VENV
echo  [3/4] 필요한 프로그램을 확인/설치합니다...
echo.
echo        * WARNING 으로 시작하는 문구가 보여도 정상입니다. 무시하세요.
echo        * 창을 닫지 말고 기다려주세요.
echo.
".venv-mcp\Scripts\python.exe" -m pip install --upgrade pip --quiet --disable-pip-version-check
".venv-mcp\Scripts\python.exe" -m pip install -r requirements-mcp.txt --disable-pip-version-check
if errorlevel 1 goto PIP_FAIL
echo.
echo        설치 완료
echo.

REM ---------- 4. Claude 데스크톱 앱에 등록 ----------
echo  [4/4] Claude 데스크톱 앱에 등록합니다...
"%~dp0.venv-mcp\Scripts\python.exe" "%~dp0mcp_setup.py" "%DARTKEY%" "%~dp0.venv-mcp\Scripts\python.exe" "%~dp0dart_mcp.py"
if errorlevel 1 goto CFG_FAIL

echo.
echo  ================================================
echo    설치가 끝났습니다.
echo.
echo    Claude 데스크톱 앱을 완전히 껐다가 다시 켜주세요.
echo    (작업표시줄 아이콘 우클릭 - 종료)
echo.
echo    다시 켜면 대화창 아래 도구 아이콘에
echo    opendart 가 보입니다.
echo  ================================================
echo.
pause
exit /b 0

:NO_KEY
echo.
echo  [중단] API 키를 입력하지 않으셨습니다.
echo         https://opendart.fss.or.kr 에서 발급받은 뒤 다시 실행해주세요.
echo.
pause
exit /b 1

:VENV_FAIL
echo.
echo  [오류] 실행 환경을 만들지 못했습니다.
echo         이 폴더가 OneDrive/바탕화면 동기화 폴더 안에 있으면
echo         C:\GPCM 같은 단순한 경로로 옮긴 뒤 다시 시도해주세요.
echo.
pause
exit /b 1

:PIP_FAIL
echo.
echo  [오류] 필요한 프로그램 설치에 실패했습니다.
echo         인터넷 연결을 확인한 뒤 다시 실행해주세요.
echo         회사망이라면 방화벽이 pypi.org 를 막고 있을 수 있습니다.
echo.
pause
exit /b 1

:CFG_FAIL
echo.
echo  [오류] Claude 데스크톱 앱 설정에 등록하지 못했습니다.
echo         위에 표시된 내용을 설정 파일에 직접 넣어주세요.
echo         Claude 데스크톱 앱이 설치되어 있는지도 확인해주세요.
echo.
pause
exit /b 1
