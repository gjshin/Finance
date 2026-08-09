@echo off
chcp 65001 >nul
title GPCM Calculator (DART) - 국내 상장사용
cd /d "%~dp0"

echo.
echo  ================================================
echo    GPCM Calculator (DART) - 국내 상장사용
echo  ================================================
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
echo  [1/3] 파이썬 확인 완료
echo.

REM ---------- 2. 실행 환경 준비 ----------
if exist ".venv\Scripts\python.exe" goto HAVE_VENV

echo  [2/3] 최초 실행입니다. 준비하는 데 2~5분 걸립니다.
echo        (다음부터는 이 과정을 건너뜁니다)
echo.
%PY% -m venv .venv
if errorlevel 1 goto VENV_FAIL

:HAVE_VENV
echo  [2/3] 필요한 프로그램을 확인/설치합니다...
echo.
".venv\Scripts\python.exe" -m pip install --upgrade pip --quiet
".venv\Scripts\python.exe" -m pip install -r requirements.txt --quiet
if errorlevel 1 goto PIP_FAIL
echo        설치 완료
echo.

REM ---------- 3. 앱 실행 ----------
echo  [3/3] 앱을 실행합니다. 잠시 후 브라우저가 자동으로 열립니다.
echo.
echo  ------------------------------------------------
echo   종료하려면 이 검은 창을 그냥 닫으세요.
echo  ------------------------------------------------
echo.
".venv\Scripts\python.exe" -m streamlit run gpcm_kr.py
pause
exit /b 0

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
