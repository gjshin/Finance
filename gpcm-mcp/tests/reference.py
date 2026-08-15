"""원본 gpcm_kr.py 의 계산 함수를 그대로 불러온다.

이 패키지는 gpcm_kr.py 를 복사해 만들었다. 복사본이 원본과 같은 숫자를 내는지
기계적으로 증명할 방법이 없으면, 두 계산기는 조용히 갈라진다. 그래서 원본을
직접 불러와 같은 입력을 먹이고 결과를 대조한다.

원본을 그냥 import 할 수는 없다.

- L34 에서 st.set_page_config 를 부른다 (streamlit 설치 필요)
- L2411 에서 KRX 상장목록을 실제로 내려받는다 (import 만 해도 네트워크에 나간다)
- 위젯 호출이 50 개 넘게 있다

그래서 `with st.sidebar:` 앞까지만 잘라서 exec 한다. 계산 함수와 export 함수는
전부 그 앞에 있고, 잘라낸 뒤쪽은 UI 뿐이다. 원본 파일은 읽기만 한다 — 이 패키지의
전제가 "gpcm_kr.py 를 고치지 않는다" 이므로 손대서는 안 된다.
"""

import sys
import types
from pathlib import Path

SIDEBAR_MARKER = "with st.sidebar:"

# tests/reference.py -> gpcm-mcp/ -> Finance/
ORIGINAL = Path(__file__).resolve().parents[2] / "gpcm_kr.py"


class _StreamlitStub(types.ModuleType):
    """원본이 import 시점에 실제로 부르는 것만 흉내낸다.

    잘라낸 앞부분이 쓰는 streamlit 기능은 셋뿐이다: set_page_config 한 번과
    cache_resource 데코레이터 두 번. cache_resource 는 `@st.cache_resource` 와
    `@st.cache_resource(ttl=3600)` 두 형태로 모두 쓰이므로 둘 다 받아야 한다.

    캐시는 일부러 구현하지 않는다. 테스트는 매번 새 가짜 데이터를 먹이는데
    캐시가 살아 있으면 첫 번째 것이 계속 나온다.
    """

    def __init__(self):
        super().__init__("streamlit")

    def set_page_config(self, *args, **kwargs):
        return None

    def cache_resource(self, *args, **kwargs):
        # @st.cache_resource  (데코레이터로 직접 사용)
        if len(args) == 1 and callable(args[0]) and not kwargs:
            return args[0]

        # @st.cache_resource(ttl=3600)  (호출 후 데코레이터 반환)
        def decorator(fn):
            return fn

        return decorator

    # 잘라낸 앞부분에는 없지만, 원본이 조금 바뀌어도 테스트가 엉뚱한 곳에서
    # 죽지 않도록 나머지 속성은 아무것도 하지 않는 함수로 돌려준다.
    def __getattr__(self, name):
        def _noop(*args, **kwargs):
            return None

        return _noop


def _compute_source():
    text = ORIGINAL.read_text(encoding="utf-8")
    idx = text.find(SIDEBAR_MARKER)
    if idx == -1:
        raise RuntimeError(
            f"{ORIGINAL} 에서 {SIDEBAR_MARKER!r} 를 찾지 못했습니다. "
            "원본 구조가 바뀌었다면 이 로더를 함께 고쳐야 합니다."
        )
    return text[:idx]


_module = None


def load():
    """원본의 계산 계층만 담은 모듈을 돌려준다. 한 번만 exec 하고 재사용한다."""
    global _module
    if _module is not None:
        return _module

    if not ORIGINAL.exists():
        raise RuntimeError(f"원본을 찾을 수 없습니다: {ORIGINAL}")

    stub = _StreamlitStub()
    saved = sys.modules.get("streamlit")
    sys.modules["streamlit"] = stub
    try:
        mod = types.ModuleType("gpcm_kr_reference")
        mod.__file__ = str(ORIGINAL)
        exec(compile(_compute_source(), str(ORIGINAL), "exec"), mod.__dict__)
    finally:
        if saved is None:
            sys.modules.pop("streamlit", None)
        else:
            sys.modules["streamlit"] = saved

    _module = mod
    return mod
