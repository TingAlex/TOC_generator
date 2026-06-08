"""打印后端：把 PDF 静默打印到 OneNote 桌面版打印机。

复刻 OneNote Batch 的「print to OneNote」路径：配合 client.set_printout_section() 先把
打印输出定向到目标分区，再由本模块把每个 PDF 打到「OneNote (Desktop)」打印机，OneNote 即在该
分区新建一页打印输出。

打印引擎用 **SumatraPDF**（免费便携、静默打印、打完自动退出、不抢焦点），比 Adobe 的 `/t` 稳。
仅 Windows + OneNote 桌面版可用。
"""

import os
import shutil
import subprocess
from pathlib import Path

# OneNote 桌面版（Office16）打印机名。注意避开 UWP 版「Send to Microsoft OneNote」
# （端口含包名 …8wekyb3d8bbwe…），本项目只支持桌面版，COM SetFilingLocation 也走桌面栈。
DESKTOP_PRINTER = "OneNote (Desktop)"

# SumatraPDF 常见安装位置（便携版也可经 --sumatra 显式指定）。
_SUMATRA_CANDIDATES = [
    r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
    r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
]


def find_sumatra(explicit: str | None = None) -> str:
    """定位 SumatraPDF.exe：优先 --sumatra 指定，其次 PATH，再次常见安装位置。

    找不到时抛 FileNotFoundError，附安装指引。
    """
    if explicit:
        p = Path(explicit)
        if not p.is_file():
            raise FileNotFoundError(f"--sumatra 指定的文件不存在：{explicit}")
        return str(p)

    on_path = shutil.which("SumatraPDF") or shutil.which("SumatraPDF.exe")
    if on_path:
        return on_path

    local = os.environ.get("LOCALAPPDATA")
    candidates = list(_SUMATRA_CANDIDATES)
    if local:
        candidates.append(str(Path(local) / "SumatraPDF" / "SumatraPDF.exe"))
    for c in candidates:
        if Path(c).is_file():
            return c

    raise FileNotFoundError(
        "未找到 SumatraPDF。请安装（https://www.sumatrapdfreader.org/download-free-pdf-viewer），"
        "或用 --sumatra 指定 SumatraPDF.exe 的完整路径。"
    )


def print_pdf(pdf_path: str | Path, printer: str, sumatra: str,
              timeout: float = 120.0) -> None:
    """用 SumatraPDF 把单个 PDF 静默打印到指定打印机（阻塞至 SumatraPDF 退出）。

    打印是异步的：SumatraPDF 退出只代表已把作业交给打印机；OneNote 何时生成页由调用方
    轮询分区页数确认（见 cli/onenote_import）。打印失败（非零退出）抛 CalledProcessError。
    """
    subprocess.run(
        [sumatra, "-print-to", printer, "-silent", str(pdf_path)],
        check=True, timeout=timeout,
    )
