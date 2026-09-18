"""toc-onenote-batch —— 可恢复的 OneNote 无人值守批处理。

任务清单把多本笔记本串行交给 ``toc-onenote-copy-online``。子命令本身会
验证目标是源页的原生副本前缀，因此进程中断后重跑可以从下一页继续。
本执行器另外提供：单实例锁、有限重试、持久日志、原子 JSON 状态和隐藏后台启动。
"""

from __future__ import annotations

import argparse
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
import hashlib
import json
import msvcrt
import os
from pathlib import Path
import re
import subprocess
import sys
import time
from typing import Callable, Iterator, TextIO


class BatchPlanError(RuntimeError):
    """任务清单、状态或锁不安全，不应继续。"""


@dataclass(frozen=True)
class BatchJob:
    source_notebook: str
    target_notebook: str | None = None
    source_section_group: str | None = None
    create_target: bool = True
    ref_notebook: str | None = None
    sync_settle: float = 8.0
    ready_timeout: float = 120.0

    @classmethod
    def from_dict(cls, raw: dict, index: int) -> "BatchJob":
        if not isinstance(raw, dict):
            raise BatchPlanError(f"第 {index} 个任务必须是 JSON 对象。")
        source = str(raw.get("source_notebook", "")).strip()
        if not source:
            raise BatchPlanError(f"第 {index} 个任务缺少 source_notebook。")
        target = raw.get("target_notebook")
        group = raw.get("source_section_group")
        reference = raw.get("ref_notebook")
        settle = float(raw.get("sync_settle", 8.0))
        timeout = float(raw.get("ready_timeout", 120.0))
        if settle < 0 or timeout <= 0:
            raise BatchPlanError(
                f"第 {index} 个任务的 sync_settle 必须 >= 0，"
                "ready_timeout 必须 > 0。")
        return cls(
            source_notebook=source,
            target_notebook=(str(target).strip() if target else None),
            source_section_group=(str(group).strip() if group else None),
            create_target=bool(raw.get("create_target", True)),
            ref_notebook=(str(reference).strip() if reference else None),
            sync_settle=settle,
            ready_timeout=timeout,
        )

    @property
    def target_name(self) -> str:
        return self.target_notebook or self.source_notebook


@dataclass(frozen=True)
class BatchPlan:
    jobs: tuple[BatchJob, ...]
    sha256: str


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def load_plan(path: Path) -> BatchPlan:
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise BatchPlanError(f"无法读取任务清单 {path}：{exc}") from exc
    if not isinstance(raw, dict) or raw.get("version") != 1:
        raise BatchPlanError("任务清单必须是 version=1 的 JSON 对象。")
    raw_jobs = raw.get("jobs")
    if not isinstance(raw_jobs, list) or not raw_jobs:
        raise BatchPlanError("任务清单的 jobs 必须是非空数组。")
    jobs = tuple(BatchJob.from_dict(item, i)
                 for i, item in enumerate(raw_jobs, 1))
    identities = [(job.source_notebook, job.target_name,
                   job.source_section_group) for job in jobs]
    if len(set(identities)) != len(identities):
        raise BatchPlanError("任务清单中存在重复的源/目标/分区组任务。")
    canonical = json.dumps(raw, ensure_ascii=False, sort_keys=True,
                           separators=(",", ":")).encode("utf-8")
    return BatchPlan(jobs=jobs, sha256=hashlib.sha256(canonical).hexdigest())


def build_command(job: BatchJob, *, write: bool) -> list[str]:
    command = [
        sys.executable, "-m", "tocgen.cli.onenote_copy_online",
        "--source-notebook", job.source_notebook,
        "--target-notebook", job.target_name,
        "--sync-settle", str(job.sync_settle),
        "--ready-timeout", str(job.ready_timeout),
    ]
    if job.source_section_group:
        command.extend(["--source-section-group", job.source_section_group])
    if job.create_target:
        command.append("--create-target")
    if job.ref_notebook:
        command.extend(["--ref-notebook", job.ref_notebook])
    if write:
        command.append("--write")
    return command


def _atomic_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(
        json.dumps(value, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    os.replace(temporary, path)


def _initial_state(plan: BatchPlan, *, mode: str) -> dict:
    return {
        "schema": 1,
        "plan_sha256": plan.sha256,
        "mode": mode,
        "status": "pending",
        "started_at": _utc_now(),
        "updated_at": _utc_now(),
        "current_job": None,
        "last_line": "",
        "jobs": [
            {
                "source_notebook": job.source_notebook,
                "target_notebook": job.target_name,
                "source_section_group": job.source_section_group,
                "status": "pending",
                "attempts": 0,
                "last_exit_code": None,
                "last_line": "",
            }
            for job in plan.jobs
        ],
    }


def _load_or_create_state(path: Path, plan: BatchPlan, *, mode: str) -> dict:
    if not path.exists():
        return _initial_state(plan, mode=mode)
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise BatchPlanError(f"状态文件无法读取：{path}：{exc}") from exc
    if state.get("plan_sha256") != plan.sha256:
        raise BatchPlanError(
            "状态文件属于另一份任务清单；请换一个 --state 路径，"
            "不要覆盖旧任务记录。")
    if state.get("mode") != mode:
        raise BatchPlanError("同一状态文件不能在 dry-run 和 write 之间复用。")
    if len(state.get("jobs", [])) != len(plan.jobs):
        raise BatchPlanError("状态文件的任务数与清单不一致。")
    return state


def lock_is_held(path: Path) -> bool:
    """探测调度锁；状态写着 running 但锁未持有即为孤儿状态。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    stream = path.open("a+b")
    try:
        stream.seek(0, os.SEEK_END)
        if stream.tell() == 0:
            stream.write(b"0")
            stream.flush()
        stream.seek(0)
        try:
            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            return True
        stream.seek(0)
        msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
        return False
    finally:
        stream.close()


@contextmanager
def single_instance_lock(path: Path) -> Iterator[None]:
    """Windows 文件区间锁；进程退出后由系统自动释放。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    stream = path.open("a+b")
    try:
        stream.seek(0, os.SEEK_END)
        if stream.tell() == 0:
            stream.write(b"0")
            stream.flush()
        stream.seek(0)
        try:
            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as exc:
            raise BatchPlanError(
                f"另一个 OneNote 批处理已在运行（锁：{path}）。") from exc
        yield
    finally:
        try:
            stream.seek(0)
            msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
        stream.close()


_COPY_PROGRESS = re.compile(r"本次\s+(\d+)/(\d+)")


def _update_from_line(state: dict, job_state: dict, line: str) -> None:
    clean = line.rstrip("\r\n")
    state["last_line"] = clean
    state["updated_at"] = _utc_now()
    job_state["last_line"] = clean
    match = _COPY_PROGRESS.search(clean)
    if match:
        job_state["copied_this_run"] = int(match.group(1))
        job_state["remaining_at_start"] = int(match.group(2))


def run_child(command: list[str], *, cwd: Path, log: TextIO,
              on_line: Callable[[str], None]) -> int:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    process = subprocess.Popen(
        command,
        cwd=str(cwd),
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    assert process.stdout is not None
    for line in process.stdout:
        sys.stdout.write(line)
        sys.stdout.flush()
        log.write(line)
        log.flush()
        on_line(line)
    return process.wait()


def run_plan(plan: BatchPlan, *, state_path: Path, log_path: Path,
             lock_path: Path, cwd: Path, write: bool, max_attempts: int,
             retry_delay: float,
             child_runner: Callable[..., int] = run_child) -> int:
    if max_attempts < 1 or retry_delay < 0:
        raise BatchPlanError("max_attempts 必须 >= 1，retry_delay 必须 >= 0。")
    mode = "write" if write else "dry-run"
    with single_instance_lock(lock_path):
        state = _load_or_create_state(state_path, plan, mode=mode)
        if state.get("status") == "complete":
            print("任务已完成；无需重复执行。")
            return 0
        if state.get("status") in {"running", "retrying"}:
            state["recoveries"] = int(state.get("recoveries", 0)) + 1
            state["recovered_at"] = _utc_now()
            for item in state.get("jobs", []):
                if item.get("status") in {"running", "retrying"}:
                    item["status"] = "pending"
        state["status"] = "running"
        state["runner_pid"] = os.getpid()
        state["updated_at"] = _utc_now()
        _atomic_json(state_path, state)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with log_path.open("a", encoding="utf-8", buffering=1) as log:
            log.write(f"\n=== batch start {_utc_now()} mode={mode} ===\n")
            for index, job in enumerate(plan.jobs):
                job_state = state["jobs"][index]
                if job_state.get("status") == "complete":
                    continue
                state["current_job"] = index + 1
                job_state["status"] = "running"
                _atomic_json(state_path, state)
                command = build_command(job, write=write)
                succeeded = False
                for attempt in range(1, max_attempts + 1):
                    job_state["attempts"] = int(job_state.get("attempts", 0)) + 1
                    job_state["attempt_in_launch"] = attempt
                    job_state["last_started_at"] = _utc_now()
                    _atomic_json(state_path, state)
                    header = (f"\n=== job {index + 1}/{len(plan.jobs)} "
                              f"attempt {attempt}/{max_attempts}: "
                              f"{job.source_notebook} -> {job.target_name} ===\n")
                    print(header, end="")
                    log.write(header)

                    def on_line(line: str) -> None:
                        _update_from_line(state, job_state, line)
                        _atomic_json(state_path, state)

                    exit_code = child_runner(
                        command, cwd=cwd, log=log, on_line=on_line)
                    job_state["last_exit_code"] = exit_code
                    job_state["last_finished_at"] = _utc_now()
                    if exit_code == 0:
                        job_state["status"] = "complete"
                        succeeded = True
                        _atomic_json(state_path, state)
                        break
                    job_state["status"] = "retrying" if attempt < max_attempts else "failed"
                    _atomic_json(state_path, state)
                    if attempt < max_attempts:
                        message = f"任务失败，{retry_delay:g} 秒后从已验证前缀重试。\n"
                        print(message, end="")
                        log.write(message)
                        log.flush()
                        time.sleep(retry_delay)
                if not succeeded:
                    state["status"] = "failed"
                    state["finished_at"] = _utc_now()
                    state["updated_at"] = _utc_now()
                    _atomic_json(state_path, state)
                    return int(job_state.get("last_exit_code") or 1)

            state["status"] = "complete"
            state["current_job"] = None
            state["finished_at"] = _utc_now()
            state["updated_at"] = _utc_now()
            _atomic_json(state_path, state)
            log.write(f"=== batch complete {_utc_now()} ===\n")
            print("全部 OneNote 批处理任务已完成。")
            return 0


def _default_sidecar(plan_path: Path, suffix: str) -> Path:
    return plan_path.with_name(plan_path.stem + suffix)


def _detach(args: argparse.Namespace) -> int:
    command = [
        sys.executable, "-m", "tocgen.cli.onenote_batch",
        "--plan", str(args.plan),
        "--state", str(args.state),
        "--log", str(args.log),
        "--lock", str(args.lock),
        "--max-attempts", str(args.max_attempts),
        "--retry-delay", str(args.retry_delay),
    ]
    if args.write:
        command.append("--write")
    creationflags = (
        subprocess.CREATE_NEW_PROCESS_GROUP
        | subprocess.DETACHED_PROCESS
        | subprocess.CREATE_NO_WINDOW
    )
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    process = subprocess.Popen(
        command,
        cwd=os.getcwd(),
        env=env,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=creationflags,
        close_fds=True,
    )
    print(f"已启动隐藏后台任务，PID={process.pid}")
    print(f"状态：{args.state}")
    print(f"日志：{args.log}")
    return 0


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(line_buffering=True)
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True,
                        help="version=1 的 JSON 任务清单")
    parser.add_argument("--state", type=Path, default=None,
                        help="持久状态 JSON（默认与 plan 同目录）")
    parser.add_argument("--log", type=Path, default=None,
                        help="追加日志（默认与 plan 同目录）")
    parser.add_argument("--lock", type=Path, default=None,
                        help="单实例锁（默认与 plan 同目录）")
    parser.add_argument("--max-attempts", type=int, default=3,
                        help="每本笔记本本次启动最多尝试次数（默认 3）")
    parser.add_argument("--retry-delay", type=float, default=30.0,
                        help="失败重试前等待秒数（默认 30）")
    parser.add_argument("--write", action="store_true",
                        help="真正复制；缺省让子任务做 dry-run")
    parser.add_argument("--detach", action="store_true",
                        help="以隐藏后台进程运行，当前命令立即返回")
    parser.add_argument("--status", action="store_true",
                        help="只输出状态 JSON 后退出")
    args = parser.parse_args()

    args.plan = args.plan.resolve()
    args.state = (args.state or _default_sidecar(args.plan, ".state.json")).resolve()
    args.log = (args.log or _default_sidecar(args.plan, ".log")).resolve()
    args.lock = (args.lock or _default_sidecar(args.plan, ".lock")).resolve()

    if args.status:
        if not args.state.exists():
            sys.exit(f"状态文件尚未生成：{args.state}")
        state = json.loads(args.state.read_text(encoding="utf-8"))
        active = lock_is_held(args.lock)
        state["process_active"] = active
        if state.get("status") in {"running", "retrying"} and not active:
            state["effective_status"] = "interrupted"
            state["diagnostic"] = (
                "状态为运行中，但单实例锁未被进程持有；可安全重启续跑。")
        else:
            state["effective_status"] = state.get("status")
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return
    try:
        plan = load_plan(args.plan)
        if args.detach:
            raise SystemExit(_detach(args))
        raise SystemExit(run_plan(
            plan,
            state_path=args.state,
            log_path=args.log,
            lock_path=args.lock,
            cwd=Path.cwd(),
            write=args.write,
            max_attempts=args.max_attempts,
            retry_delay=args.retry_delay,
        ))
    except BatchPlanError as exc:
        sys.exit(f"✗ {exc}")


if __name__ == "__main__":
    main()
