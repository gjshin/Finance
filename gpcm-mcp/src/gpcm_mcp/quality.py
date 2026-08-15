"""수집 실패 지점을 모으는 기록장. gpcm_kr.py 에서 그대로 옮겼다."""

from .constants import SEV_ERROR, SEV_INFO, SEV_WARN  # noqa: F401  (재수출)


class QualityLog:
    """자동 수집이 실패하거나 값을 채우지 못한 지점을 모은다.

    DART 조회는 회사·기간별로 조용히 실패한다. 지금까지는 실패한 자리에 0이 남고
    나머지가 계속 돌아, 엑셀만 봐서는 그 0이 '정말 0'인지 '못 가져온 것'인지
    구분할 수 없었다. 특히 LTM은 (당기누계 + 전기연간 - 전년동기)라서 셋 중
    하나만 빠져도 숫자가 그럴듯하게 틀린다.
    """

    def __init__(self):
        self.rows = []

    def add(self, level, ticker, company, item, message):
        self.rows.append({
            'Level': level, 'Ticker': ticker, 'Company': company,
            'Item': item, 'Message': message,
        })

    def has(self, level):
        return any(r['Level'] == level for r in self.rows)
