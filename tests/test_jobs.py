"""Version 4 background jobs: the build must keep running when the user changes page, and every end must be reported.
Synthetic data only.   python -m unittest discover -s tests -v"""
import os
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))
os.environ.setdefault("INVOICE_DB_PATH", str(Path(tempfile.mkdtemp()) / "jobs.sqlite3"))
os.environ.setdefault("INVOICE_AUDIT_LOG", str(Path(tempfile.mkdtemp()) / "audit.log"))

from app_core import jobs, runner  # noqa: E402


def wait(m, owner, timeout=20):
    """Wait until the owner has no active job."""
    end = time.time() + timeout
    while time.time() < end:
        if m.active(owner) is None:
            return
        time.sleep(0.02)
    raise AssertionError("job did not end")


class JobTests(unittest.TestCase):
    def setUp(self):
        self.m = jobs.JobManager(max_workers=2)

    def tearDown(self):
        self.m._pool.shutdown(wait=False, cancel_futures=True)

    def test_success_result_is_collected_once(self):
        def work(progress):
            progress(0.5, "half")
            return {"answer": 42}
        self.m.submit("a", "build", "Build", "x.xlsx", work)
        wait(self.m, "a")
        (view, run), = self.m.collect("a")
        self.assertEqual(view.status, jobs.DONE)
        self.assertEqual(run, {"answer": 42})
        self.assertEqual(view.fraction, 1.0)
        self.assertEqual(self.m.collect("a"), [])             # handed over once; no second copy kept

    def test_keeps_running_while_nobody_looks(self):
        """The page-switch case: the job must finish although nothing polls it."""
        go = threading.Event()
        self.m.submit("a", "build", "Build", "x.xlsx", lambda p: (go.wait(5), "ok")[1])
        time.sleep(0.1)
        self.assertEqual(self.m.active("a").status, jobs.RUNNING)
        go.set()
        wait(self.m, "a")
        self.assertEqual(self.m.collect("a")[0][1], "ok")

    def test_failure_is_reported_not_lost(self):
        def work(progress):
            progress(0.3, "x")
            raise ValueError("bad workbook")
        self.m.submit("a", "build", "Build", "x.xlsx", work)
        wait(self.m, "a")
        (view, run), = self.m.collect("a")
        self.assertEqual((view.status, view.error, run), (jobs.FAILED, "bad workbook", None))
        self.assertIn(view.id, [v.id for v in self.m.jobs("a")])   # still listed for the Dashboard until dismissed

    def test_invalid_master_keeps_every_problem(self):
        def work(progress):
            raise runner.InvalidMaster(["Row 5: QTY missing", "Row 9: bad HTS"])
        self.m.submit("a", "build", "Build", "x.xlsx", work)
        wait(self.m, "a")
        v = self.m.jobs("a")[0]
        self.assertEqual(v.status, jobs.FAILED)
        self.assertEqual(v.problems, ("Row 5: QTY missing", "Row 9: bad HTS"))

    def test_base_exception_does_not_kill_the_worker(self):
        def work(progress):
            raise SystemExit("boom")
        self.m.submit("a", "build", "Build", "x.xlsx", work)
        wait(self.m, "a")
        self.assertEqual(self.m.jobs("a")[0].status, jobs.FAILED)
        self.m.submit("a", "build", "Build", "y.xlsx", lambda p: 1)          # the pool still works
        wait(self.m, "a")

    def test_cancel_running_job(self):
        started = threading.Event()

        def work(progress):
            started.set()
            for i in range(500):
                time.sleep(0.01)
                progress(i / 500, "step")                 # the cancel point
            return "finished anyway"
        job = self.m.submit("a", "build", "Build", "x.xlsx", work)
        started.wait(2)
        self.assertTrue(self.m.cancel("a", job.id))
        wait(self.m, "a")
        (view, run), = self.m.collect("a")
        self.assertEqual((view.status, run), (jobs.CANCELLED, None))

    def test_cancel_swallowed_by_broad_except_still_cancels(self):
        started = threading.Event()

        def work(progress):
            started.set()
            for _ in range(500):
                time.sleep(0.01)
                try:
                    progress(0.1, "x")
                except Exception:                         # a careless `except Exception` in the work
                    pass
            return "finished anyway"
        job = self.m.submit("a", "build", "Build", "x.xlsx", work)
        started.wait(2)
        self.m.cancel("a", job.id)
        wait(self.m, "a")
        self.assertEqual(self.m.jobs("a")[0].status, jobs.CANCELLED)

    def test_one_active_job_per_session(self):
        go = threading.Event()
        self.m.submit("a", "build", "Build", "x.xlsx", lambda p: go.wait(5))
        with self.assertRaises(jobs.JobBusy):
            self.m.submit("a", "build", "Build", "y.xlsx", lambda p: 1)
        self.m.submit("b", "build", "Build", "z.xlsx", lambda p: 1)            # another session is not blocked
        go.set()
        wait(self.m, "a")
        self.m.submit("a", "build", "Build", "y.xlsx", lambda p: 1)            # free again after the end

    def test_sessions_are_isolated(self):
        go = threading.Event()
        job = self.m.submit("a", "build", "Build", "secret.xlsx", lambda p: (go.wait(5), "A's invoices")[1])
        self.assertEqual(self.m.jobs("b"), [])
        self.assertIsNone(self.m.active("b"))
        self.assertFalse(self.m.cancel("b", job.id))                        # b cannot cancel a's job by guessing the id
        self.m.dismiss("b", job.id)
        go.set()
        wait(self.m, "a")
        self.assertEqual(self.m.collect("b"), [])                           # nor take its result
        self.assertEqual(self.m.collect("a")[0][1], "A's invoices")

    def test_waits_for_a_worker_then_runs(self):
        m = jobs.JobManager(max_workers=1)
        gate = threading.Event()
        try:
            m.submit("a", "build", "Build", "1.xlsx", lambda p: gate.wait(5))
            second = m.submit("b", "build", "Build", "2.xlsx", lambda p: "two")
            time.sleep(0.1)
            self.assertEqual(m.active("b").status, jobs.QUEUED)
            gate.set()
            wait(m, "b")
            self.assertEqual(m.collect("b")[0][1], "two")
            self.assertEqual(second.status, jobs.QUEUED)                    # the view returned by submit is a snapshot
        finally:
            gate.set()
            m._pool.shutdown(wait=False, cancel_futures=True)

    def test_cancel_while_queued_never_runs(self):
        m = jobs.JobManager(max_workers=1)
        gate, ran = threading.Event(), []
        try:
            m.submit("a", "build", "Build", "1.xlsx", lambda p: gate.wait(5))
            job = m.submit("b", "build", "Build", "2.xlsx", lambda p: ran.append(1))
            self.assertTrue(m.cancel("b", job.id))
            self.assertEqual(m.jobs("b")[0].status, jobs.CANCELLED)
            gate.set()
            wait(m, "a")
            time.sleep(0.1)
            self.assertEqual(ran, [])
        finally:
            gate.set()
            m._pool.shutdown(wait=False, cancel_futures=True)

    def test_discard_owner_on_sign_out(self):
        started = threading.Event()

        def work(progress):
            started.set()
            for _ in range(500):
                time.sleep(0.01)
                progress(0.1, "x")
        self.m.submit("a", "build", "Build", "x.xlsx", work)
        self.m.submit("b", "build", "Build", "y.xlsx", lambda p: 1)
        started.wait(2)
        self.m.discard_owner("a")
        self.assertEqual(self.m.jobs("a"), [])
        wait(self.m, "b")
        self.assertEqual(len(self.m.jobs("b")), 1)                          # others untouched

    def test_dismiss_and_cleanup(self):
        m = jobs.JobManager(max_workers=1, ttl=0.05, keep=2)
        try:
            for i in range(4):
                m.submit("a", "build", "Build", f"{i}.xlsx", lambda p: 1)
                wait(m, "a")
            self.assertLessEqual(len(m.jobs("a")), 2)                       # only the last `keep` are kept
            time.sleep(0.1)
            self.assertEqual(m.jobs("a"), [])                               # and they expire
        finally:
            m._pool.shutdown(wait=False)

    def test_dismiss_marks_failure(self):
        def work(progress):
            raise ValueError("x")
        job = self.m.submit("a", "build", "Build", "x.xlsx", work)
        wait(self.m, "a")
        self.assertFalse(self.m.jobs("a")[0].dismissed)
        self.m.dismiss("a", job.id)
        self.assertTrue(self.m.jobs("a")[0].dismissed)

    def test_singleton(self):
        self.assertIs(jobs.manager(), jobs.manager())


class RealBuildTests(unittest.TestCase):
    """The real runner inside a worker thread, with the work running while 'the page' (this thread) does something else."""

    def setUp(self):
        self.m = jobs.JobManager()

    def tearDown(self):
        self.m._pool.shutdown(wait=False, cancel_futures=True)

    def test_real_build_in_background(self):
        from test_core import make_master
        data = make_master()
        seen = []
        self.m.submit("a", "build", "Build", "MASTER_INVOICE.xlsx",
                      lambda progress: runner.build_invoices(data, "MASTER_INVOICE.xlsx", progress=progress))
        while self.m.active("a"):
            v = self.m.active("a")
            if v:
                seen.append(v.fraction)
            time.sleep(0.01)
        (view, run), = self.m.collect("a")
        self.assertEqual(view.status, jobs.DONE, view.error)
        self.assertEqual(run["result"]["totals"]["invoices"], 1)
        self.assertTrue(run["files"] and run["report_pdf"])
        self.assertEqual(sorted(seen), seen)                                # progress only moves forward

    def test_real_bad_master_reports_problems(self):
        self.m.submit("a", "build", "Build", "bad.xlsx", lambda progress: runner.build_invoices(b"not a workbook", "bad.xlsx", progress=progress))
        wait(self.m, "a")
        v = self.m.jobs("a")[0]
        self.assertEqual(v.status, jobs.FAILED)
        self.assertTrue(v.problems)

    def test_real_rebuild_in_background(self):
        from test_core import make_master
        run0 = runner.build_invoices(make_master(), "MASTER_INVOICE.xlsx")
        self.m.submit("a", "rebuild", "Rebuild", "MASTER_INVOICE.xlsx", lambda progress: runner.rebuild(run0, progress=progress))
        wait(self.m, "a")
        (view, run), = self.m.collect("a")
        self.assertEqual(view.status, jobs.DONE, view.error)
        self.assertEqual(run["run_id"], run0["run_id"])
        self.assertTrue(run["result"]["rebuilt"])


if __name__ == "__main__":
    unittest.main()
