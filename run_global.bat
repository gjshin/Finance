@echo off
chcp 65001 >nul
title GPCM Calculator - 해외 상장사용
cd /d "%~dp0"

echo.
echo  ================================================
echo    GPCM Calculator - 해외 상장사용
echo  ================================================
echo.

REM 임시 파일을 %TEMP% 대신 여기에 만든다. 회사 PC의 보안 프로그램이
REM %TEMP% 에 새 실행 파일 만드는 걸 막는 경우가 있어서다.
set "GPCM_HOME=%LOCALAPPDATA%\gpcm"
if not exist "%GPCM_HOME%\tmp" mkdir "%GPCM_HOME%\tmp" >nul 2>&1
set "TMP=%GPCM_HOME%\tmp"
set "TEMP=%GPCM_HOME%\tmp"

REM ---------- 1. uv 찾기 ----------
REM uv가 파이썬까지 알아서 준비한다. 파이썬을 따로 설치할 필요가 없다.
set "UV=%USERPROFILE%\.local\bin\uv.exe"
if exist "%UV%" goto HAVE_UV

for /f "delims=" %%i in ('where uv 2^>nul') do set "UV=%%i"
if exist "%UV%" goto HAVE_UV

echo  [1/2] 처음 실행이라 준비 도구를 설치합니다 (1분 내외)...
echo.
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex" >nul 2>&1
set "UV=%USERPROFILE%\.local\bin\uv.exe"
if not exist "%UV%" goto UV_FAIL

:HAVE_UV
echo  [1/2] 준비 도구 확인 완료
echo.

REM ---------- 2. 앱 실행 ----------
echo  [2/2] 앱을 실행합니다.
echo.
echo        * 처음에는 파이썬과 필요한 프로그램을 받느라 3~5분 걸립니다.
echo          창을 닫지 말고 기다려주세요. 다음부터는 바로 뜹니다.
echo        * WARNING 으로 시작하는 문구가 보여도 정상입니다.
echo.
echo  ------------------------------------------------
echo   준비가 끝나면 브라우저가 자동으로 열립니다.
echo   종료하려면 이 검은 창을 그냥 닫으세요.
echo  ------------------------------------------------
echo.

"%UV%" run --no-project --python 3.12 --with-requirements requirements.txt python -m streamlit run GPCM.py
if errorlevel 1 goto RUN_FAIL
pause
exit /b 0

:UV_FAIL
echo.
echo  [오류] 준비 도구(uv)를 설치하지 못했습니다.
echo.
echo         회사망이라면 방화벽이 astral.sh 를 막고 있을 수 있습니다.
echo         전산팀에 문의하시거나 다른 네트워크에서 한 번 실행해보세요.
echo.
pause
exit /b 1

:RUN_FAIL
echo.
echo  [오류] 앱을 실행하지 못했습니다.
echo.
echo         - 인터넷 연결을 확인해주세요
echo         - 회사망이라면 방화벽이 pypi.org 를 막고 있을 수 있습니다
echo         - 위에 빨간 글씨가 있으면 그 부분을 캡처해서 문의해주세요
echo.
pause
exit /b 1
