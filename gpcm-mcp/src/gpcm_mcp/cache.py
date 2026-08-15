"""수명이 있는 캐시.

원본은 @st.cache_resource(ttl=3600) 으로 KRX 상장목록을 한 시간 동안 재사용했다.
lru_cache 는 기계적으로는 바꿔 끼울 수 있지만 만료가 없다. Streamlit 앱은 오래 떠 있어도
사용자가 새로고침하면 세션이 새로 뜨는데, MCP 서버는 며칠씩 같은 프로세스로 산다.
그 사이 신규 상장된 종목이 조용히 '없는 회사'가 되는 것을 막으려면 TTL 이 필요하다.
"""

import time


def ttl_cached(seconds):
    """인자 없는 로더용 TTL 캐시. 만료 전이면 같은 객체를 돌려준다."""

    def decorator(fn):
        box = {}

        def wrapper():
            now = time.time()
            if 'value' in box and now - box['at'] < seconds:
                return box['value']
            value = fn()
            box['value'] = value
            box['at'] = now
            return value

        def clear():
            box.clear()

        wrapper.cache_clear = clear
        wrapper.__name__ = fn.__name__
        wrapper.__doc__ = fn.__doc__
        return wrapper

    return decorator
