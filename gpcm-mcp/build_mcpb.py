"""gpcm-kr.mcpb 를 만든다.

.mcpb 는 클로드 데스크톱 앱이 더블클릭 한 번으로 설치하는 꾸러미다. 설치하면
다른 커넥터들과 같은 목록에 뜨고 거기서 끄고 켤 수 있다. 설정 파일을 손으로
고칠 일도, 경로를 적을 일도 없다.

서버 타입은 "uv" 를 쓴다. 전통적인 방식은 의존성을 server/lib 에 넣어 두는 것인데,
pandas·numpy·openpyxl 처럼 **컴파일된 패키지는 그렇게 담을 수 없다** (OS·파이썬
버전마다 다른 바이너리가 필요하다). uv 타입은 pyproject.toml 만 넣어 두면 설치하는
쪽에서 파이썬과 의존성을 알아서 준비한다.

    python build_mcpb.py

만들어진 gpcm-kr.mcpb 를 더블클릭하면 설치된다.
"""

import json
import shutil
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
STAGE = HERE / 'build' / 'mcpb'
OUT = HERE / 'gpcm-kr.mcpb'

VERSION = '0.1.0'

# 앱이 파이썬과 의존성을 준비할 때 쓰는 파일. 버전은 브라우저 앱(requirements.txt)과
# 같은 조합으로 고정한다 — 같은 라이브러리로 같은 숫자가 나와야 한다.
PYPROJECT = f'''[project]
name = "gpcm-mcp"
version = "{VERSION}"
description = "국내 상장사 GPCM 배수·WACC 계산 (DART 기반)"
requires-python = ">=3.11"
dependencies = [
    "mcp==2.0.0",
    "pandas==3.0.5",
    "numpy==2.4.6",
    "requests==2.33.1",
    "openpyxl==3.1.5",
    "OpenDartReader==0.2.2",
    "finance-datareader==0.9.202",
    "yfinance==1.5.2",
    "lxml==6.1.1",
]

[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[tool.hatch.build.targets.wheel]
packages = ["src/gpcm_mcp"]
'''

# uv 가 실행하는 진입점. 패키지 안의 server.py 는 상대 import 를 쓰므로 스크립트로
# 직접 돌릴 수 없다. 얇은 껍데기를 하나 두고 거기서 부른다.
ENTRY = '''"""gpcm-kr 진입점. 실제 내용은 gpcm_mcp/server.py 에 있다."""

from gpcm_mcp.server import main

if __name__ == "__main__":
    main()
'''

MANIFEST = {
    'manifest_version': '0.4',
    'name': 'gpcm-kr',
    'display_name': 'GPCM 계산기 (국내 상장사)',
    'version': VERSION,
    'description': 'DART 공시로 국내 상장사 GPCM 배수와 WACC 를 계산해 엑셀로 만듭니다.',
    'long_description': (
        '유사기업 배수(GPCM)와 WACC 를 계산해 조서용 엑셀로 뽑습니다. '
        'DART 재무제표와 KRX 주가를 직접 조회하며, 계산식은 기존 Streamlit 앱'
        '(gpcm_kr.py)에서 그대로 옮겨온 것이라 같은 숫자가 나옵니다.\n\n'
        '**국내에서만 동작합니다.** DART 가 해외 접속을 차단합니다.\n\n'
        'DART 인증키는 https://opendart.fss.or.kr 에서 무료로 발급받습니다. '
        '하루 조회 한도가 키마다 주어지므로 각자 본인 키를 쓰십시오.'
    ),
    'author': {'name': 'SGJ'},
    'keywords': ['dart', 'gpcm', 'valuation', 'wacc', 'korea', '밸류에이션'],
    'server': {
        'type': 'uv',
        'entry_point': 'src/server.py',
        'mcp_config': {
            'env': {
                'OPENDART_API_KEY': '${user_config.api_key}',
                'GPCM_MCP_OUTPUT_DIR': '${user_config.output_dir}',
            },
        },
    },
    'user_config': {
        'api_key': {
            'type': 'string',
            'title': 'OpenDART 인증키',
            'description': ('https://opendart.fss.or.kr 에서 무료로 발급받은 40자 키. '
                            '각자 본인 키를 쓰십시오 — 하루 조회 한도가 키마다 주어집니다.'),
            'sensitive': True,
            'required': True,
        },
        'output_dir': {
            'type': 'directory',
            'title': '엑셀 저장 폴더',
            'description': '계산 결과 엑셀을 저장할 곳입니다.',
            'required': False,
            'default': '${DOCUMENTS}',
        },
    },
    'tools': [
        {'name': 'check_dart_access', 'description': 'DART 도달 여부와 키 설정 확인'},
        {'name': 'latest_filed_period', 'description': '공시가 끝난 최신 분기'},
        {'name': 'list_peer_candidates', 'description': 'KRX 업종·비교대상 후보 조회'},
        {'name': 'gpcm_valuation', 'description': 'GPCM 배수와 WACC 계산 (억원)'},
        {'name': 'historical_financials', 'description': '다기간 재무제표 요약 (백만원)'},
        {'name': 'gpcm_job_status', 'description': '조회 진행 상황과 결과'},
        {'name': 'gpcm_job_cancel', 'description': '조회 중단'},
    ],
    'compatibility': {'platforms': ['win32', 'darwin', 'linux']},
}


def build():
    if STAGE.exists():
        shutil.rmtree(STAGE)
    (STAGE / 'src').mkdir(parents=True)

    shutil.copytree(HERE / 'src' / 'gpcm_mcp', STAGE / 'src' / 'gpcm_mcp',
                    ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
    (STAGE / 'src' / 'server.py').write_text(ENTRY, encoding='utf-8')
    (STAGE / 'pyproject.toml').write_text(PYPROJECT, encoding='utf-8')
    (STAGE / 'manifest.json').write_text(
        json.dumps(MANIFEST, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')

    if OUT.exists():
        OUT.unlink()
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        for path in sorted(STAGE.rglob('*')):
            if path.is_file():
                z.write(path, path.relative_to(STAGE).as_posix())

    size_kb = OUT.stat().st_size / 1024
    print(f'{OUT.name}  ({size_kb:.0f} KB)')
    with zipfile.ZipFile(OUT) as z:
        names = z.namelist()
    print(f'  파일 {len(names)}개')
    for n in ('manifest.json', 'pyproject.toml', 'src/server.py',
              'src/gpcm_mcp/server.py'):
        print(f'  {"있음" if n in names else "없음!!"}  {n}')
    return OUT


if __name__ == '__main__':
    build()
