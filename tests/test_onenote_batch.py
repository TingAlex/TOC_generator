import json
from pathlib import Path
import tempfile
import unittest

from tocgen.cli.onenote_batch import (
    BatchPlanError,
    build_command,
    load_plan,
    run_plan,
)


class OneNoteBatchTests(unittest.TestCase):
    def write_plan(self, root: Path, jobs: list[dict]) -> Path:
        path = root / "plan.json"
        path.write_text(
            json.dumps({"version": 1, "jobs": jobs}, ensure_ascii=False),
            encoding="utf-8",
        )
        return path

    def paths(self, root: Path) -> tuple[Path, Path, Path]:
        return root / "state.json", root / "batch.log", root / "batch.lock"

    def test_load_plan_and_build_child_command(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            plan = load_plan(self.write_plan(root, [{
                "source_notebook": "本地书",
                "target_notebook": "在线书",
                "source_section_group": "正文",
                "create_target": True,
                "sync_settle": 9,
                "ready_timeout": 150,
            }]))

            command = build_command(plan.jobs[0], write=True)

            self.assertIn("tocgen.cli.onenote_copy_online", command)
            self.assertIn("本地书", command)
            self.assertIn("在线书", command)
            self.assertIn("正文", command)
            self.assertIn("--create-target", command)
            self.assertEqual(command[-1], "--write")

    def test_failure_retries_and_progress_is_persisted(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            plan = load_plan(self.write_plan(root, [{
                "source_notebook": "本地书",
                "target_notebook": "在线书",
            }]))
            state_path, log_path, lock_path = self.paths(root)
            calls = []

            def fake_runner(command, *, cwd, log, on_line):
                calls.append(command)
                on_line("✓ 第三页（本次 3/10）\n")
                return 1 if len(calls) == 1 else 0

            result = run_plan(
                plan,
                state_path=state_path,
                log_path=log_path,
                lock_path=lock_path,
                cwd=root,
                write=True,
                max_attempts=2,
                retry_delay=0,
                child_runner=fake_runner,
            )
            state = json.loads(state_path.read_text(encoding="utf-8"))

            self.assertEqual(result, 0)
            self.assertEqual(len(calls), 2)
            self.assertEqual(state["status"], "complete")
            self.assertEqual(state["jobs"][0]["status"], "complete")
            self.assertEqual(state["jobs"][0]["attempts"], 2)
            self.assertEqual(state["jobs"][0]["copied_this_run"], 3)

    def test_completed_state_is_not_run_again(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            plan = load_plan(self.write_plan(root, [{
                "source_notebook": "本地书",
            }]))
            state_path, log_path, lock_path = self.paths(root)

            def success(command, *, cwd, log, on_line):
                return 0

            self.assertEqual(run_plan(
                plan,
                state_path=state_path,
                log_path=log_path,
                lock_path=lock_path,
                cwd=root,
                write=True,
                max_attempts=1,
                retry_delay=0,
                child_runner=success,
            ), 0)

            def must_not_run(*args, **kwargs):
                raise AssertionError("完成的任务不应再次执行")

            self.assertEqual(run_plan(
                plan,
                state_path=state_path,
                log_path=log_path,
                lock_path=lock_path,
                cwd=root,
                write=True,
                max_attempts=1,
                retry_delay=0,
                child_runner=must_not_run,
            ), 0)

    def test_changed_plan_cannot_reuse_old_state(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            first = load_plan(self.write_plan(root, [{
                "source_notebook": "第一本",
            }]))
            state_path, log_path, lock_path = self.paths(root)

            def success(command, *, cwd, log, on_line):
                return 0

            run_plan(
                first,
                state_path=state_path,
                log_path=log_path,
                lock_path=lock_path,
                cwd=root,
                write=True,
                max_attempts=1,
                retry_delay=0,
                child_runner=success,
            )
            second = load_plan(self.write_plan(root, [{
                "source_notebook": "第二本",
            }]))

            with self.assertRaisesRegex(BatchPlanError, "另一份任务清单"):
                run_plan(
                    second,
                    state_path=state_path,
                    log_path=log_path,
                    lock_path=lock_path,
                    cwd=root,
                    write=True,
                    max_attempts=1,
                    retry_delay=0,
                    child_runner=success,
                )

    def test_orphan_running_state_is_recovered_on_restart(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            plan = load_plan(self.write_plan(root, [{
                "source_notebook": "可续跑书",
            }]))
            state_path, log_path, lock_path = self.paths(root)

            def crash(command, *, cwd, log, on_line):
                raise KeyboardInterrupt("模拟调度进程被关闭")

            with self.assertRaises(KeyboardInterrupt):
                run_plan(
                    plan,
                    state_path=state_path,
                    log_path=log_path,
                    lock_path=lock_path,
                    cwd=root,
                    write=True,
                    max_attempts=1,
                    retry_delay=0,
                    child_runner=crash,
                )

            def success(command, *, cwd, log, on_line):
                return 0

            self.assertEqual(run_plan(
                plan,
                state_path=state_path,
                log_path=log_path,
                lock_path=lock_path,
                cwd=root,
                write=True,
                max_attempts=1,
                retry_delay=0,
                child_runner=success,
            ), 0)
            state = json.loads(state_path.read_text(encoding="utf-8"))
            self.assertEqual(state["recoveries"], 1)
            self.assertEqual(state["status"], "complete")


if __name__ == "__main__":
    unittest.main()
