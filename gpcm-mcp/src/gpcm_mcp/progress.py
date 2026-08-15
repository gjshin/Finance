"""진행 상황 보고.

원본은 status_container(st.status)와 progress_bar(st.progress)를 계산 함수에
인자로 넘긴다. 그 자리에 끼울 물건을 만든다.

일부러 같은 오리 타이핑 모양(write/update/progress)을 유지했다. 인터페이스를
새로 설계하면 이식한 함수 본문 아홉 곳을 고쳐야 하는데, 고치지 않은 줄은
숫자를 바꿀 수 없다는 게 이 패키지의 전부다.

덤으로 취소도 여기서 해결된다. 계산 루프는 이미 반복마다 progress() 를 부르므로,
취소 요청이 있을 때 그 자리에서 예외를 던지면 본문을 건드리지 않고 멈출 수 있다.
"""


class JobCancelled(Exception):
    """사용자가 작업을 취소했다."""


class NullProgress:
    """아무것도 보고하지 않는다. 테스트와 동기 호출용."""

    def write(self, *args, **kwargs):
        pass

    def update(self, *args, **kwargs):
        pass

    def progress(self, *args, **kwargs):
        pass


class JobProgress:
    """작업 레코드에 진행 상황을 기록한다.

    fetch_financial_data 는 progress((idx)/total) 로 0-based 비율을 넘긴다.
    즉 첫 회사에서 0.0 이 오고 루프 안에서는 1.0 에 닿지 않는다 — 끝나고 나서
    따로 progress(1.0) 을 부른다. 그 성질을 그대로 두고 비율만 옮긴다.
    """

    def __init__(self, job):
        self._job = job

    def _check(self):
        if self._job.cancel_requested:
            raise JobCancelled()

    def write(self, text='', *args, **kwargs):
        self._check()
        self._job.note(str(text))

    def update(self, *args, label=None, **kwargs):
        self._check()
        if label:
            self._job.stage = str(label)
            self._job.note(str(label))

    def progress(self, value=0.0, *args, **kwargs):
        self._check()
        try:
            self._job.fraction = max(0.0, min(1.0, float(value)))
        except (TypeError, ValueError):
            pass
