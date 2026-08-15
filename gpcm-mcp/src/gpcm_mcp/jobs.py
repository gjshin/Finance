"""오래 걸리는 조회를 작업으로 관리한다.

회사 수 × 기간 수만큼 DART 를 두드린다. 10 개사 × 8 분기면 수백 번이고 몇 분이
걸린다. 도구 호출을 그동안 붙잡고 있으면 타임아웃에 걸린다.

그래서 작업을 띄우고 잠깐 기다린다. 그 안에 끝나면 결과를 바로 돌려주고
(2~3 개사 조회는 보통 여기서 끝난다), 안 끝나면 job_id 를 주고 물러난다.
작은 조회는 한 번에 끝나고 큰 조회는 전송을 막지 않는다.

작업자를 하나만 둔 것은 의도다. 계산 코드가 회사마다 time.sleep(0.5) 로
DART 호출 간격을 벌리고 있는데, 작업 둘이 동시에 돌면 그 간격이 반으로 준다.
"""

import threading
import traceback
import uuid
from collections import deque
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

from .progress import JobCancelled

MAX_KEPT = 20


class Job:
    def __init__(self, kind, label):
        self.id = f'{kind}_{uuid.uuid4().hex[:10]}'
        self.kind = kind
        self.label = label
        self.status = 'queued'          # queued | running | done | failed | cancelled
        self.stage = ''
        self.fraction = 0.0
        self.result = None
        self.error = None
        self.cancel_requested = False
        self.started_at = None
        self.finished_at = None
        self.created_at = datetime.now()
        self._log = deque(maxlen=30)
        self.done_event = threading.Event()

    def note(self, text):
        if text:
            self._log.append(text)

    @property
    def elapsed_s(self):
        if self.started_at is None:
            return 0.0
        end = self.finished_at or datetime.now()
        return round((end - self.started_at).total_seconds(), 1)

    def snapshot(self):
        out = {
            'job_id': self.id,
            'status': self.status,
            'label': self.label,
            'stage': self.stage,
            'progress_pct': round(self.fraction * 100, 1),
            'elapsed_s': self.elapsed_s,
            'log_tail': list(self._log)[-5:],
        }
        if self.status == 'failed':
            out['error'] = self.error
        return out


class JobStore:
    def __init__(self):
        self._jobs = {}
        self._order = deque()
        self._lock = threading.Lock()
        self._pool = ThreadPoolExecutor(max_workers=1,
                                        thread_name_prefix='gpcm')

    def submit(self, kind, label, fn):
        job = Job(kind, label)
        with self._lock:
            self._jobs[job.id] = job
            self._order.append(job.id)
            self._evict()

        def run():
            job.status = 'running'
            job.started_at = datetime.now()
            try:
                job.result = fn(job)
                job.status = 'done'
                job.fraction = 1.0
            except JobCancelled:
                job.status = 'cancelled'
            except Exception as exc:
                job.status = 'failed'
                job.error = f'{type(exc).__name__}: {exc}'
                job.note(traceback.format_exc(limit=3))
            finally:
                job.finished_at = datetime.now()
                job.done_event.set()

        self._pool.submit(run)
        return job

    def _evict(self):
        """끝난 작업만 오래된 것부터 버린다. 파일은 지우지 않는다."""
        while len(self._order) > MAX_KEPT:
            for job_id in list(self._order):
                job = self._jobs.get(job_id)
                if job and job.status in ('done', 'failed', 'cancelled'):
                    self._order.remove(job_id)
                    self._jobs.pop(job_id, None)
                    break
            else:
                return

    def get(self, job_id):
        with self._lock:
            return self._jobs.get(job_id)

    def cancel(self, job_id):
        job = self.get(job_id)
        if job is None:
            return None
        if job.status in ('done', 'failed', 'cancelled'):
            return job
        job.cancel_requested = True
        return job


STORE = JobStore()
