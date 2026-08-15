"""산출물 저장.

엑셀 파일이 결과물이다. MCP 도구는 글자만 돌려줄 수 있으므로, 파일을 디스크에
쓰고 경로를 알려준다. 워크북을 base64 로 실어 보내는 건 하지 않는다 — 8 기간
10 개사면 수백 KB 라 대화 맥락만 잡아먹고 얻는 게 없다.
"""

import os
import re
from datetime import datetime
from pathlib import Path

ENV_DIR = 'GPCM_MCP_OUTPUT_DIR'
DEFAULT_DIR = Path.home() / 'Documents' / 'GPCM_Reports'

_UNSAFE = re.compile(r'[^0-9A-Za-z가-힣._-]+')


def output_dir():
    """저장 폴더. 없으면 만든다.

    기본값을 문서 폴더로 잡은 이유: 이 서버는 각자 본인 PC 에서 돌고,
    파일 탐색기로 바로 찾아갈 수 있어야 한다.
    """
    raw = os.environ.get(ENV_DIR, '').strip()
    path = Path(raw).expanduser() if raw else DEFAULT_DIR
    path.mkdir(parents=True, exist_ok=True)
    return path


def _safe_stem(text):
    return _UNSAFE.sub('_', str(text)).strip('_') or 'report'


def build_name(prefix, *parts):
    """덮어쓰지 않도록 시각을 붙인 파일명. 앱의 명명 규칙을 잇는다."""
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    body = '_'.join(_safe_stem(p) for p in parts if p not in (None, ''))
    return f"{prefix}_{body}_{stamp}.xlsx" if body else f"{prefix}_{stamp}.xlsx"


def save(book, filename):
    """BytesIO 를 파일로 쓴다. 중간에 죽어도 반쪽짜리 파일이 남지 않게 한다.

    반쪽 워크북은 아예 없는 것보다 나쁘다 — 열리기는 하는데 숫자가 빠져 있다.
    """
    directory = output_dir()
    final = directory / filename
    tmp = directory / (filename + '.tmp')
    data = book.getvalue() if hasattr(book, 'getvalue') else book
    tmp.write_bytes(data)
    os.replace(tmp, final)
    return final
