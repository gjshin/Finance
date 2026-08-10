@echo off
chcp 65001 >nul
title DART 접속 진단
cd /d "%~dp0"

echo.
echo  ================================================
echo    DART 접속 진단
echo  ================================================
echo.
echo   이 PC(한국)에서 DART에 연결되는지 확인합니다.
echo   나온 결과를 복사해두셨다가,
echo   Streamlit Cloud에서 같은 진단을 돌린 결과와 비교합니다.
echo.

REM ---------- 파이썬 찾기 ----------
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
REM 앱과 같은 가상환경을 쓴다. 없으면 만든다.
if exist ".venv\Scripts\python.exe" goto HAVE_VENV

echo  최초 실행입니다. 준비하는 데 2~5분 걸립니다.
echo.
%PY% -m venv .venv
if errorlevel 1 goto VENV_FAIL
".venv\Scripts\python.exe" -m pip install --upgrade pip --quiet --disable-pip-version-check
".venv\Scripts\python.exe" -m pip install -r requirements.txt --disable-pip-version-check
if errorlevel 1 goto PIP_FAIL

:HAVE_VENV
echo  진단 화면을 엽니다. 브라우저에서 "진단 시작"을 누르세요.
echo.
echo  ------------------------------------------------
echo   끝나면 이 검은 창을 그냥 닫으세요.
echo  ------------------------------------------------
echo.
".venv\Scripts\python.exe" -m streamlit run diagnose_dart.py
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
echo.
pause
exit /b 1
