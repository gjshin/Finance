"""단위 테스트 — 기간 구성, 시트명, 세율 수식, 패키지 경계."""

import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parents[1] / 'src'))

from gpcm_mcp.excel.gpcm_book import _tax_rate_formula  # noqa: E402
from gpcm_mcp.excel.sheetnames import (quote_for_formula,  # noqa: E402
                                       sanitize_sheet_name, unique_sheet_names)
from gpcm_mcp.orphans import build_periods_to_fetch, build_target_periods  # noqa: E402

SRC = Path(__file__).parents[1] / 'src' / 'gpcm_mcp'


# --- 기간 구성 --------------------------------------------------------------

def test_quarterly_periods_span_years():
    assert build_target_periods(2024, '3Q', 2025, '2Q') == [
        '2024.3Q', '2024.4Q', '2025.1Q', '2025.2Q']


def test_annual_periods_use_4q_only():
    assert build_target_periods(2023, '1Q', 2025, '2Q', cycle='annual') == [
        '2023.4Q', '2024.4Q', '2025.4Q']


def test_single_period_is_allowed():
    assert build_target_periods(2025, '1Q', 2025, '1Q') == ['2025.1Q']


def test_reversed_range_is_empty_not_crash():
    """빈 목록이 나와야 한다. 원본은 이걸 그대로 흘려보내 ZeroDivisionError 를 냈다."""
    assert build_target_periods(2025, '1Q', 2024, '4Q') == []
    assert build_target_periods(2025, '3Q', 2025, '1Q') == []


def test_historical_annual_labels():
    assert build_periods_to_fetch(2023, 2024) == [
        {'year': 2023, 'qtr': None, 'label': '2023년'},
        {'year': 2024, 'qtr': None, 'label': '2024년'}]


def test_historical_quarterly_labels():
    got = build_periods_to_fetch(2024, 2024, '2Q', '4Q')
    assert [p['label'] for p in got] == ['2024년 2Q', '2024년 3Q', '2024년 4Q']


# --- 시트명 ----------------------------------------------------------------

@pytest.mark.parametrize('raw,expected', [
    ('삼성전자', '삼성전자'),
    ('한국ABC/DEF', '한국ABC_DEF'),
    ('브라켓[주]', '브라켓_주_'),
    ('물음표?주식회사', '물음표_주식회사'),
    ('별표*컴퍼니', '별표_컴퍼니'),
    ('콜론:테스트', '콜론_테스트'),
    ('역슬래시\\테스트', '역슬래시_테스트'),
])
def test_illegal_characters_replaced(raw, expected):
    assert sanitize_sheet_name(raw) == expected


def test_name_truncated_to_31():
    assert len(sanitize_sheet_name('가' * 50)) == 31


def test_empty_name_gets_placeholder():
    assert sanitize_sheet_name('') == 'Sheet'
    assert sanitize_sheet_name(None) == 'Sheet'


def test_same_company_name_maps_to_one_sheet():
    """이름이 완전히 같으면 같은 회사다. 시트도 하나여야 한다 (원본과 같은 동작)."""
    got = unique_sheet_names(['같은이름', '같은이름'])
    assert got == {'같은이름': '같은이름'}


def test_different_names_never_share_a_sheet():
    got = unique_sheet_names(['에이[주]', '에이_주_'])
    assert got['에이[주]'] != got['에이_주_']
    assert len(set(got.values())) == 2


def test_duplicates_only_after_truncation():
    """앞 31자가 같은 두 회사 — 원본은 여기서 시트명이 겹쳐 깨졌다."""
    a, b = '가' * 31 + 'AAA', '가' * 31 + 'BBB'
    got = unique_sheet_names([a, b])
    assert got[a] != got[b]
    assert all(len(v) <= 31 for v in got.values())


def test_illegal_chars_can_also_collide():
    got = unique_sheet_names(['A/B', 'A:B'])
    assert got['A/B'] != got['A:B']


def test_formula_quoting_doubles_apostrophes():
    assert quote_for_formula("O'Brien") == "O''Brien"


# --- 세율 수식 (원본과 의도적으로 다른 유일한 곳) ---------------------------

def test_tax_formula_matches_original_for_fy2025():
    """FY2025 이하에서는 원본과 글자까지 같아야 한다."""
    original = '=IF(AD6<=2, 0.099, IF(AD6<=200, 0.209, IF(AD6<=3000, 0.231, 0.264)))'
    assert _tax_rate_formula('AD6', '2025.1Q') == original
    assert _tax_rate_formula('AD6', '2024.4Q') == original
    assert _tax_rate_formula('AD6', '2023.2Q') == original


def test_tax_formula_uses_2026_brackets_from_fy2026():
    """원본은 여기서도 FY2023~25 세율을 써서 파이썬 WACC 와 모순됐다."""
    got = _tax_rate_formula('AD6', '2026.2Q')
    assert '0.11' in got and '0.22' in got and '0.242' in got and '0.275' in got
    assert '0.099' not in got


def test_tax_formula_matches_python_brackets():
    """엑셀 수식과 파이썬 계산이 같은 구간표를 써야 한다."""
    from gpcm_mcp.tax import get_korean_marginal_tax_rate
    for period, pretax, expected in [
        ('2025.1Q', 1, 0.099), ('2025.1Q', 100, 0.209),
        ('2025.1Q', 1000, 0.231), ('2025.1Q', 5000, 0.264),
        ('2026.2Q', 1, 0.110), ('2026.2Q', 100, 0.220),
        ('2026.2Q', 1000, 0.242), ('2026.2Q', 5000, 0.275),
    ]:
        year = int(period.split('.')[0])
        assert get_korean_marginal_tax_rate(pretax, year) == expected
        assert str(expected).rstrip('0') in _tax_rate_formula('AD6', period)


# --- 패키지 경계 ------------------------------------------------------------

def test_no_streamlit_anywhere():
    hits = []
    for path in SRC.rglob('*.py'):
        for i, line in enumerate(path.read_text(encoding='utf-8').splitlines(), 1):
            stripped = line.strip()
            if stripped.startswith(('import streamlit', 'from streamlit')):
                hits.append(f'{path.name}:{i}')
    assert not hits, f'streamlit 을 쓰는 곳이 남았다: {hits}'


def test_openpyxl_only_under_excel():
    hits = []
    for path in SRC.glob('*.py'):
        text = path.read_text(encoding='utf-8')
        if 'import openpyxl' in text or 'from openpyxl' in text:
            hits.append(path.name)
    assert not hits, f'openpyxl 은 excel/ 안에서만 써야 한다: {hits}'


def test_mcp_only_in_server():
    hits = []
    for path in SRC.rglob('*.py'):
        if path.name == 'server.py':
            continue
        text = path.read_text(encoding='utf-8')
        if 'from mcp' in text or 'import mcp' in text:
            hits.append(path.name)
    assert not hits, f'mcp 는 server.py 에서만 써야 한다: {hits}'


def test_import_does_not_touch_the_network():
    """원본은 import 만 해도 KRX 상장목록을 내려받았다. 여기서는 그러면 안 된다."""
    code = (
        "import socket\n"
        "def blocked(*a, **k):\n"
        "    raise AssertionError('import 중에 네트워크로 나갔다')\n"
        "socket.socket.connect = blocked\n"
        "socket.create_connection = blocked\n"
        "import gpcm_mcp, gpcm_mcp.gpcm, gpcm_mcp.historical, gpcm_mcp.runner\n"
        "import gpcm_mcp.excel.gpcm_book, gpcm_mcp.excel.historical_book\n"
        "print('OK')\n"
    )
    proc = subprocess.run([sys.executable, '-c', code], capture_output=True,
                          text=True, cwd=str(SRC.parent))
    assert proc.returncode == 0, proc.stderr[-2000:]
    assert 'OK' in proc.stdout
