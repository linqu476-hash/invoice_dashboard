"""Background job engine (version 4). No Streamlit in here, so it can be tested on its own.

WHY: Streamlit re-runs the page script on every click and STOPS the running script when the user opens another page. A build that
ran inside the Build page therefore died when the user clicked Dashboard (and nothing was reported). Now the build runs in a worker
THREAD of the server process. Pages only start it, watch it and collect the result, so switching pages, re-runs and even closing the
browser tab no longer stop it.

Rules:
  * one build at a time per browser session (JobBusy); at most MAX_WORKERS builds run at once on the server, the others wait (queued)
  * every operation needs the owner (a random id kept in the session): a session can never see, cancel or collect another one's job
  * a build can be cancelled: the progress callback checks a flag and stops the work at the next step
  * every end state is kept and reported: done, failed (with the reason and, for a bad master, every problem with its row) or cancelled
  * finished jobs are forgotten after TTL_SECONDS, and each session keeps only its last KEEP_PER_OWNER jobs, so memory cannot grow
  * jobs live in server memory: a server restart / redeploy ends them (they show as lost when the page is opened again, not as running)
"""
import logging
import threading
import time
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime, timezone

from . import security

QUEUED, RUNNING, DONE, FAILED, CANCELLED = "queued", "running", "done", "failed", "cancelled"
ACTIVE = (QUEUED, RUNNING)
MAX_WORKERS = 2
TTL_SECONDS = 2 * 3600
KEEP_PER_OWNER = 10

_log = logging.getLogger("app_core.jobs")


class JobBusy(Exception):
    """The session already has a build that is queued or running."""


class JobCancelled(BaseException):
    """Raised inside the worker by the progress callback after a cancel request.
    BaseException on purpose: a broad `except Exception` in the work must not swallow the cancel."""


@dataclass(frozen=True)
class JobView:
    """Read-only copy of a job for the pages. Holds no file bytes."""
    id: str
    kind: str
    label: str
    filename: str
    status: str
    fraction: float
    text: str
    created: str
    started: str | None
    finished: str | None
    elapsed: float
    error: str
    problems: tuple
    cancelling: bool
    dismissed: bool

    @property
    def active(self):
        return self.status in ACTIVE


class _Job:
    def __init__(self, owner, kind, label, filename):
        self.id = uuid.uuid4().hex
        self.owner, self.kind, self.label, self.filename = owner, kind, label, filename
        self.status, self.fraction, self.text = QUEUED, 0.0, "Waiting for a free worker ..."
        self.created = _now()
        self.started = self.finished = None
        self.t0 = self.t1 = None                    # monotonic clock
        self.error, self.problems = "", ()
        self.run = None                             # the result until a page collects it
        self.collected = False
        self.dismissed = False
        self.cancel = threading.Event()
        self.future = None
        self.finished_at = 0.0                      # wall clock, for the TTL


def _now():
    return datetime.now(timezone.utc).isoformat()


class JobManager:
    def __init__(self, max_workers=MAX_WORKERS, ttl=TTL_SECONDS, keep=KEEP_PER_OWNER):
        self._pool = ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="invoice-job")
        self._jobs = {}
        self._lock = threading.RLock()
        self._ttl, self._keep = ttl, keep

    # ------------------------------------------------------------------ start
    def submit(self, owner, kind, label, filename, work):
        """Queue work(progress) -> result. progress(fraction, text) reports and is the cancel point. Returns a JobView."""
        if not owner:
            raise ValueError("a job needs an owner")
        with self._lock:
            self._purge()
            if any(j.owner == owner and j.status in ACTIVE for j in self._jobs.values()):
                raise JobBusy("A build is already running in this session.")
            job = _Job(owner, kind, label, filename)
            self._jobs[job.id] = job
            job.future = self._pool.submit(self._run, job, work)
            security.audit("job_queued", job=job.id[:8], kind=kind, file=filename)
            return self._view(job)

    def _run(self, job, work):
        with self._lock:
            if job.cancel.is_set():                  # cancelled while it was waiting
                self._end(job, CANCELLED)
                return
            job.status, job.text = RUNNING, "Starting ..."
            job.started, job.t0 = _now(), time.monotonic()
        security.audit("job_started", job=job.id[:8], kind=job.kind)

        def progress(fraction, text):
            if job.cancel.is_set():
                raise JobCancelled()
            with self._lock:
                job.fraction = max(0.0, min(float(fraction), 1.0))
                job.text = str(text)

        try:
            result = work(progress)
            if job.cancel.is_set():                  # cancel asked during the very last step: honour it, drop the result
                raise JobCancelled()
        except JobCancelled:
            self._end(job, CANCELLED)
        except BaseException as e:                   # noqa: BLE001  nothing may kill the worker silently
            problems = tuple(str(p) for p in (getattr(e, "problems", None) or ()))
            if problems:                             # a bad master is a user error, not a crash: no stack trace
                _log.warning("job %s: the master has %d problem(s)", job.id[:8], len(problems))
            else:
                _log.error("job %s failed:\n%s", job.id[:8], traceback.format_exc())
            self._end(job, FAILED, error=("The master workbook has problems." if problems else (str(e) or e.__class__.__name__))[:500],
                      problems=problems)
        else:
            self._end(job, DONE, result=result)

    def _end(self, job, status, result=None, error="", problems=()):
        with self._lock:
            job.status, job.run, job.error, job.problems = status, result, error, problems
            job.t1, job.finished, job.finished_at = time.monotonic(), _now(), time.time()
            if status == DONE:
                job.fraction, job.text = 1.0, "Done"
            elif status == CANCELLED:
                job.text = "Cancelled"
            seconds = round((job.t1 - job.t0), 2) if job.t0 else 0
        security.audit(f"job_{status}", job=job.id[:8], kind=job.kind, seconds=seconds, **({"error": error} if error else {}))

    # ------------------------------------------------------------------ look / control (always scoped to the owner)
    def jobs(self, owner):
        """This session's jobs, newest first."""
        with self._lock:
            self._purge()
            return [self._view(j) for j in sorted(self._owned(owner), key=lambda j: j.created, reverse=True)]

    def active(self, owner):
        return next((j for j in self.jobs(owner) if j.active), None)

    def cancel(self, owner, job_id):
        with self._lock:
            job = self._get(owner, job_id)
            if job and job.status in ACTIVE:
                job.cancel.set()
                if job.future and job.future.cancel():       # still waiting: it will never start
                    self._end(job, CANCELLED)
                else:
                    job.text = "Cancelling ..."
                return True
        return False

    def collect(self, owner):
        """Jobs of this session that ended since the last call, as [(JobView, run | None)]. A result is handed over ONCE
        (the manager does not keep a second copy of the files)."""
        out = []
        with self._lock:
            for j in sorted(self._owned(owner), key=lambda j: j.finished_at):
                if j.status not in ACTIVE and not j.collected:
                    j.collected = True
                    run, j.run = j.run, None
                    out.append((self._view(j), run if j.status == DONE else None))
        return out

    def dismiss(self, owner, job_id):
        with self._lock:
            job = self._get(owner, job_id)
            if job and job.status not in ACTIVE:
                job.dismissed = True

    def discard_owner(self, owner):
        """Sign-out / expired session: stop the owner's work and forget everything about it."""
        if not owner:
            return
        with self._lock:
            for j in list(self._owned(owner)):
                j.cancel.set()
                if j.future:
                    j.future.cancel()
                del self._jobs[j.id]

    # ------------------------------------------------------------------ internals
    def _owned(self, owner):
        return [j for j in self._jobs.values() if j.owner == owner]

    def _get(self, owner, job_id):
        j = self._jobs.get(job_id)
        return j if j is not None and j.owner == owner else None        # another session's id is simply "not found"

    def _view(self, j):
        now = time.monotonic()
        elapsed = ((j.t1 or now) - j.t0) if j.t0 else 0.0
        return JobView(j.id, j.kind, j.label, j.filename, j.status, j.fraction, j.text, j.created, j.started, j.finished, elapsed,
                       j.error, j.problems, j.cancel.is_set() and j.status in ACTIVE, j.dismissed)

    def _purge(self):
        now = time.time()
        for j in [j for j in self._jobs.values() if j.status not in ACTIVE and now - j.finished_at > self._ttl]:
            del self._jobs[j.id]
        by_owner = {}
        for j in self._jobs.values():
            by_owner.setdefault(j.owner, []).append(j)
        for js in by_owner.values():
            for j in sorted(js, key=lambda j: j.created, reverse=True)[self._keep:]:
                if j.status not in ACTIVE:
                    del self._jobs[j.id]


_MANAGER = None
_GUARD = threading.Lock()


def manager():
    """The one JobManager of the server process (module-level, so it survives Streamlit re-runs and page switches)."""
    global _MANAGER
    with _GUARD:
        if _MANAGER is None:
            _MANAGER = JobManager()
        return _MANAGER
