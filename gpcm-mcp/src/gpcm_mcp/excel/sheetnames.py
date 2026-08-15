"""회사명을 엑셀 시트명으로 바꾼다.

원본은 회사명을 `comp[:31]` 로 자르기만 해서 두 가지로 깨진다.

1. 엑셀 시트명에 쓸 수 없는 글자 `[ ] : * ? / \` 가 이름에 있으면 openpyxl 이
   예외를 낸다. `SK C&C` 는 괜찮지만 `한국ABC/DEF` 같은 이름은 그대로 죽는다.
2. 앞 31 자가 같은 회사가 둘이면 같은 시트명이 나온다.

고칠 때 함정이 하나 있다. 잘린 이름은 시트를 만들 때(historical_book L124)만
쓰이는 게 아니라 Summary 시트의 SUMIFS 수식 문자열(L77, L95)에도 박힌다.
시트 만드는 쪽만 고치면 **수식이 없는 시트를 가리키게 되어** Summary 가
#REF! 또는 0 이 된다. 오류도 안 나고 숫자만 조용히 틀린다.

그래서 매핑을 한 번만 만들어 양쪽이 같은 것을 쓰게 한다.
"""

import re

# 엑셀이 시트명에 허용하지 않는 글자
ILLEGAL = re.compile(r'[\[\]:*?/\\]')
MAX_LEN = 31


def sanitize_sheet_name(name):
    """엑셀이 받아들이는 시트명으로 바꾼다 (중복은 보지 않는다)."""
    text = ILLEGAL.sub('_', str(name or '').strip())
    text = text.strip("'")
    text = text[:MAX_LEN].strip()
    return text or 'Sheet'


def unique_sheet_names(names):
    """{원래 이름: 시트명} 매핑. 중복되면 뒤에 _2, _3 을 붙인다.

    엑셀은 시트명 대소문자를 구분하지 않으므로 소문자로 비교한다.
    """
    mapping = {}
    used = set()
    for name in names:
        if name in mapping:
            # 같은 이름이 두 번 나오면 같은 회사다. 시트도 하나여야 한다.
            continue
        base = sanitize_sheet_name(name)
        candidate = base
        n = 2
        while candidate.lower() in used:
            suffix = f'_{n}'
            candidate = base[:MAX_LEN - len(suffix)] + suffix
            n += 1
        used.add(candidate.lower())
        mapping[name] = candidate
    return mapping


def quote_for_formula(sheet_name):
    """수식 안에 넣을 시트명. 작은따옴표는 두 번 써서 escape 한다.

    원본이 `'{name}'!A:A` 형태로 이미 따옴표를 두르므로, 이름 안의 따옴표만
    처리하면 된다.
    """
    return str(sheet_name).replace("'", "''")
