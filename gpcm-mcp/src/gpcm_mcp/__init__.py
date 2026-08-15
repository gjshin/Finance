"""국내 상장사 GPCM 배수·WACC 계산.

gpcm_kr.py(Streamlit 앱)의 계산 계층을 그대로 옮긴 것이다. 같은 입력에 같은 숫자를
내는 것이 이 패키지의 전제이므로, 계산식은 원본이 바뀔 때만 함께 바꾼다.
tests/test_parity.py 가 원본과 대조해 그 전제를 지킨다.

import 시점에 네트워크를 타지 않는다. 원본은 import 만 해도 KRX 목록을 내려받았다.
"""

__version__ = "0.1.0"
