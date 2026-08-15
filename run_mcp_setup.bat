@echo off
chcp 65001 >nul
title GPCM 계산 도구 설치 (Claude 연결용)
cd /d "%~dp0"

echo.
echo  ================================================
echo    GPCM 계산 도구 설치 (Claude 연결용)
echo  ================================================
echo.
echo   설치가 끝나면 Claude 에게 이렇게 시킬 수 있습니다.
echo.
echo     "005930, 000660 으로 2026년 2분기 기준 GPCM 돌려줘"
echo     "코스닥 반도체 업종에서 피어 후보 보여줘"
echo     "이 회사들 최근 3개년 재무제표 한번에 정리해줘"
echo.
echo   브라우저로 쓰던 run_kr.bat 은 그대로 두고 씁니다.
echo   이건 추가로 설치하는 것이지 대체하는 게 아닙니다.
echo.
echo  ------------------------------------------------
echo.

REM 임시 파일을 %TEMP% 대신 여기에 만든다. 회사 PC의 보안 프로그램이
REM %TEMP% 에 새 실행 파일 만드는 걸 막는 경우가 있어서다. (run_kr.bat 과 동일)
set "GPCM_HOME=%LOCALAPPDATA%\gpcm"
if not exist "%GPCM_HOME%\tmp" mkdir "%GPCM_HOME%\tmp" >nul 2>&1
set "TMP=%GPCM_HOME%\tmp"
set "TEMP=%GPCM_HOME%\tmp"

REM ---------- 1. uv 찾기 ----------
REM run_kr.bat 을 한 번이라도 돌렸으면 이미 있다. 없으면 여기서 받는다.
set "UV=%USERPROFILE%\.local\bin\uv.exe"
if exist "%UV%" goto HAVE_UV

for /f "delims=" %%i in ('where uv 2^>nul') do set "UV=%%i"
if exist "%UV%" goto HAVE_UV

echo  [1/4] 준비 도구를 설치합니다 (1분 내외)...
echo.
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex" >nul 2>&1
set "UV=%USERPROFILE%\.local\bin\uv.exe"
if not exist "%UV%" goto UV_FAIL

:HAVE_UV
echo  [1/4] 준비 도구 확인 완료
echo.

REM ---------- 2. API 키 ----------
echo  [2/4] OpenDART API 키가 필요합니다.
echo.
echo        브라우저 앱에 넣으시던 그 키입니다.
echo        없으시면 https://opendart.fss.or.kr 에서 무료로 발급받으세요.
echo.
echo        각자 본인 키를 쓰셔야 합니다. 하루 조회 한도가 키마다 주어져서,
echo        한 키를 여럿이 나눠 쓰면 금방 한도에 걸립니다.
echo.
set "DARTKEY="
set /p DARTKEY="        키를 붙여넣고 Enter (건너뛰려면 그냥 Enter): "
if not defined DARTKEY goto SKIP_KEY

REM setx 는 이 창이 아니라 앞으로 열릴 창에 적용된다. 그래서 아래에서 다시 안내한다.
setx OPENDART_API_KEY "%DARTKEY%" >nul 2>&1
echo.
echo        키를 저장했습니다.
goto KEY_DONE

:SKIP_KEY
echo.
echo        건너뛰었습니다. 나중에 넣으시려면 이 파일을 다시 실행하세요.

:KEY_DONE
echo.

REM ---------- 3. 설치 ----------
REM 브라우저 앱과 다른 환경에 넣는다. mcp 패키지가 요구하는 starlette 버전이
REM streamlit 과 충돌해서, 같이 설치하면 브라우저 앱이 깨진다.
echo  [3/4] 계산 도구를 설치합니다.
echo.
echo        처음에는 2~5분 걸립니다. 창을 닫지 마세요.
echo        브라우저 앱과 다른 공간에 설치하므로 앱에는 영향이 없습니다.
echo.

cd gpcm-mcp
if not exist ".venv-mcp\Scripts\python.exe" (
    "%UV%" venv .venv-mcp --python 3.12
    if errorlevel 1 goto INSTALL_FAIL
)
"%UV%" pip install --python .venv-mcp -e . --quiet
if errorlevel 1 goto INSTALL_FAIL
cd ..

echo  [3/4] 설치 완료
echo.

REM ---------- 4. Claude 에 연결 ----------
echo  [4/4] Claude 에 연결합니다.
echo.

REM JSON 에 역슬래시를 넣으면 이스케이프가 필요하다. 슬래시로 적어도 윈도우에서 동작한다.
set "PYPATH=%~dp0gpcm-mcp\.venv-mcp\Scripts\python.exe"
set "PYPATH=%PYPATH:\=/%"

REM %~dp0 는 끝에 역슬래시가 붙는다. 따옴표 바로 앞의 역슬래시는 따옴표를 먹는
REM 경우가 있어서, 안내문에 넣기 전에 떼어낸다.
set "INSTALLDIR=%~dp0"
if "%INSTALLDIR:~-1%"=="\" set "INSTALLDIR=%INSTALLDIR:~0,-1%"

> ".mcp.json" echo {
>>".mcp.json" echo   "mcpServers": {
>>".mcp.json" echo     "gpcm-kr": {
>>".mcp.json" echo       "command": "%PYPATH%",
>>".mcp.json" echo       "args": ["-m", "gpcm_mcp.server"]
>>".mcp.json" echo     }
>>".mcp.json" echo   }
>>".mcp.json" echo }

echo        .mcp.json 을 만들었습니다.
echo.
echo  ================================================
echo    설치가 끝났습니다
echo  ================================================
echo.
echo   설치된 곳 ^(이 경로를 적어두세요^):
echo.
echo       %INSTALLDIR%
echo.
echo  ------------------------------------------------
echo   [중요 1] 창을 껐다 켜야 합니다.
echo  ------------------------------------------------
echo.
echo   지금 열려 있는 창들은 방금 저장한 키를 모릅니다.
echo   이 창을 닫고, 명령 프롬프트를 새로 여세요.
echo.
echo   키가 저장됐는지는 새 창에서 이렇게 확인합니다.
echo.
echo       echo %%OPENDART_API_KEY%%
echo.
echo   키가 찍히면 정상입니다.
echo.
echo  ------------------------------------------------
echo   [중요 2] 이 PC 에서 도는 Claude 여야 합니다.
echo  ------------------------------------------------
echo.
echo   방금 설치한 것은 이 PC 안에만 있습니다.
echo   브라우저의 claude.ai 나 클라우드에서 도는 Claude 는
echo   이 PC 를 볼 수 없어서 도구를 찾지 못합니다.
echo.
echo   명령 프롬프트를 새로 열고 아래 두 줄을 그대로 실행하세요.
echo.
echo       cd /d "%INSTALLDIR%"
echo       claude
echo.
echo   Claude 가 어디서 도는지 헷갈리면 이렇게 물어보세요.
echo   "지금 작업 폴더가 어디야?"
echo   C:\ 로 시작하면 이 PC 입니다. /home/ 으로 시작하면 클라우드라 안 됩니다.
echo.
echo  ------------------------------------------------
echo.
echo   처음 한 번 "gpcm-kr 서버를 쓰겠냐" 고 물으면 허용하세요.
echo   그다음 "DART 접속 확인해줘" 라고 시켜보시면 됩니다.
echo.
echo   엑셀은 문서 폴더의 GPCM_Reports 안에 저장됩니다.
echo.
echo   [주의] 국내에서만 됩니다. DART 가 해외 접속을 막습니다.
echo.
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

:INSTALL_FAIL
echo.
echo  [오류] 계산 도구를 설치하지 못했습니다.
echo.
echo         - 인터넷 연결을 확인해주세요
echo         - 회사망이라면 방화벽이 pypi.org 를 막고 있을 수 있습니다
echo         - 위에 빨간 글씨가 있으면 그 부분을 캡처해서 문의해주세요
echo.
pause
exit /b 1
