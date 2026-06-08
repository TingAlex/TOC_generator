r"""修复 OneNote「很抱歉，OneNote 正在清理上次打开之后的内容」卡死。

复刻 OneFix（Office OneNote Gem「Fix One」）的「Fix Relaunch」：**强杀残留的 OneNote 进程
再重启**，让它免重启电脑地完成清理。**绝不删除任何缓存/数据文件**。

为什么不删缓存：本项目用户全是**在线笔记本 + 关闭同步 + 纯离线编辑**，未同步的改动只存在本地
`%LOCALAPPDATA%\Microsoft\OneNote\16.0\cache`，删它=丢未同步笔记。网上「删 16.0 缓存」的通用
解法对这种场景是危险的，故本模块只做无损的「杀进程 + 重启 + 等就绪」。

打印触发该卡死的成因：打印副本若没处理完（如导入超时），OneNote 下次启动会「清理上次内容」
并卡住；干净重启即可让它完成清理。仅 Windows + OneNote 桌面版（Office16/365 均登记为 16.0）。
"""

import subprocess
import time
from pathlib import Path

# 主程序 + 「发送到 OneNote」托盘/打印处理工具。365 与 2016 同为 Office16，进程名一致。
ONENOTE_PROCESSES = ["ONENOTE.EXE", "ONENOTEM.EXE"]

# 重启用的可执行文件常见位置（Office16 = 2016/365 点击即用与 MSI 安装）。
_EXE_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\ONENOTE.EXE",
    r"C:\Program Files\Microsoft Office\Office16\ONENOTE.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\ONENOTE.EXE",
]


def kill_onenote() -> list[str]:
    """强杀所有 OneNote 相关进程。返回实际杀掉的进程名列表。"""
    killed = []
    for image in ONENOTE_PROCESSES:
        r = subprocess.run(["taskkill", "/F", "/IM", image],
                           capture_output=True, text=True)
        if r.returncode == 0:  # 0 = 至少杀掉一个；128 = 没有该进程
            killed.append(image)
    return killed


def find_onenote_exe() -> str | None:
    """定位 ONENOTE.EXE（用于重启）。找不到则返回 None（可改由 COM 按需拉起）。"""
    for c in _EXE_CANDIDATES:
        if Path(c).is_file():
            return c
    r = subprocess.run(["where", "ONENOTE.EXE"], capture_output=True, text=True)
    if r.returncode == 0:
        first = r.stdout.strip().splitlines()
        if first:
            return first[0].strip()
    return None


def launch_onenote(exe: str | None = None) -> bool:
    """启动 OneNote（不阻塞）。返回是否成功发起启动。"""
    exe = exe or find_onenote_exe()
    if not exe:
        return False
    subprocess.Popen([exe])
    return True


def wait_until_ready(timeout: float = 120.0, poll: float = 2.0) -> bool:
    """轮询直到 OneNote COM 可响应（= 已过「清理」阶段）。返回是否就绪。"""
    from .client import OneNoteClient
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            OneNoteClient().get_hierarchy()
            return True
        except Exception:
            time.sleep(poll)
    return False


def fix_relaunch(*, settle: float = 3.0, ready_timeout: float = 120.0,
                 verbose: bool = True) -> bool:
    """杀进程 → 等待 → 重启 → 等 COM 就绪。无损（不删任何文件）。返回是否最终就绪。"""
    def log(s: str) -> None:
        if verbose:
            print(s)

    killed = kill_onenote()
    log(f"  已强杀：{', '.join(killed) if killed else '（无残留 OneNote 进程）'}")
    time.sleep(settle)

    exe = find_onenote_exe()
    if launch_onenote(exe):
        log(f"  已重启：{exe}")
    else:
        log("  未找到 ONENOTE.EXE，改由 COM 按需拉起。")

    log(f"  等待 OneNote 清理完成并恢复响应（≤{ready_timeout:.0f}s）…")
    ok = wait_until_ready(ready_timeout)
    log("  ✓ OneNote 已就绪。" if ok else "  ⚠ 超时仍未就绪——可能清理较慢或卡死较深，请稍后重试或手动查看。")
    return ok
