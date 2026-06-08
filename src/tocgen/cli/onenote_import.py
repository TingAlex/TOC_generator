"""toc-onenote-import —— Pipeline 3：把拆分 PDF 打印进 OneNote 对应分区（复刻 OneNote Batch）。

对每个分区，先用 COM SetFilingLocation 把「print to OneNote」打印输出定向到该分区，再用
SumatraPDF 把文件夹内每个 PDF 静默打到「OneNote (Desktop)」打印机；**串行**打印——打一份、
轮询分区页数 +1 确认落地、再打下一份，从而保证「打印顺序 = 页显示顺序」，事后 toc-onenote-titles
才能按顺序对齐改标题。分区需先用 toc-onenote-sections（Pipeline 2.5）建好。纯本地离线（COM）。

    toc-onenote-import --section-group "书名" --section-prefix= --root books-done/书名_拆分            # dry-run
    toc-onenote-import --section-group "书名" --section-prefix= --root books-done/书名_拆分 --write     # 正式打印
    toc-onenote-import --list --section-group "书名"                                                   # 只读看分区/页

分区 ⇄ 文件夹：去前缀后的编号 N ⇄ 文件夹 0N（配合 Pipeline 2.5 的 01…0N 用 --section-prefix=）。
"""

import argparse
import subprocess
import sys
import time
from pathlib import Path

from ..onenote.client import OneNoteClient, Section
from ..onenote.common import DEFAULT_NOTEBOOK, section_number, sorted_pdfs, resolve_scope
from ..onenote.printer import find_sumatra, print_pdf, DESKTOP_PRINTER
from ..onenote.fix import fix_relaunch

DEFAULT_SECTION_PREFIX = "新分区"
POLL_INTERVAL = 1.0  # 轮询分区页数的间隔（秒）


def _sp(s: str) -> None:
    """容错打印（控制台编码不支持的字符不致崩溃）。"""
    enc = getattr(sys.stdout, "encoding", "utf-8") or "utf-8"
    try:
        print(s)
    except UnicodeEncodeError:
        print(s.encode(enc, errors="replace").decode(enc))


def cmd_list(client: OneNoteClient, notebook_name: str, section_group: str | None) -> None:
    notebooks = client.get_hierarchy()
    nb = client.find_notebook(notebooks, notebook_name)
    if not nb:
        print(f"未找到笔记本「{notebook_name}」。可用：{', '.join(n.name for n in notebooks)}")
        return
    scope = resolve_scope(client, nb, section_group)
    if scope is None:
        return
    sections, desc = scope
    print(f"笔记本「{nb.name}」{desc}：")
    for sec in sections:
        _sp(f"\n  [{sec.name}]（{len(sec.pages)} 页）")
        for i, pg in enumerate(sec.pages, 1):
            _sp(f"    {i:>3}. {pg.name.strip() or '（空白无标题）'}")


def _wait_landed(client: OneNoteClient, section_id: str, before: int,
                 timeout: float) -> bool:
    """轮询分区页数，直到 > before（打印输出落地）或超时。返回是否落地。"""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        time.sleep(POLL_INTERVAL)
        if len(client.list_section_pages(section_id)) > before:
            return True
    return False


def print_section(client: OneNoteClient, sec: Section, folder: Path, *,
                  printer: str, sumatra: str, settle: float, timeout: float) -> dict:
    """把一个文件夹的 PDF 串行打印进对应分区。返回统计。"""
    pdfs = sorted_pdfs(folder)
    _sp(f"\n[{sec.name}] ⇄ {folder.name}/　{len(pdfs)} 个文件")
    stats = {"printed": 0, "error": 0}
    if not pdfs:
        print("    （无 PDF，跳过）")
        return stats

    # 定向：本分区成为打印输出落点
    client.set_printout_section(sec.id)

    for pdf in pdfs:
        before = len(client.list_section_pages(sec.id))
        try:
            print_pdf(pdf, printer, sumatra)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as e:
            _sp(f"    ✗ 打印失败：{pdf.name}（{e}）—— 中止本分区以防错位。")
            stats["error"] += 1
            break
        if _wait_landed(client, sec.id, before, timeout):
            stats["printed"] += 1
            _sp(f"    ✓ {pdf.name}")
            time.sleep(settle)
        else:
            _sp(f"    ⚠ 超时未见落地：{pdf.name}（{timeout:.0f}s）—— 中止本分区以防错位。")
            stats["error"] += 1
            break

    _sp(f"    小结：打印 {stats['printed']}　异常 {stats['error']}")
    return stats


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--notebook", default=DEFAULT_NOTEBOOK)
    parser.add_argument("--root", type=Path, default=None,
                        help="拆分文件夹根目录（如 books-done/书名_拆分）")
    parser.add_argument("--section-prefix", default=DEFAULT_SECTION_PREFIX,
                        help="分区名前缀（默认 新分区；配合 Pipeline 2.5 的 01…0N 用 --section-prefix=）")
    parser.add_argument("--section-group", default=None,
                        help="只处理该分区组内的分区（避免多本书的同名 0N 分区混淆）")
    parser.add_argument("--printer", default=DESKTOP_PRINTER,
                        help=f"OneNote 桌面版打印机名（默认 {DESKTOP_PRINTER}；勿用 UWP 版）")
    parser.add_argument("--sumatra", default=None,
                        help="SumatraPDF.exe 路径（缺省自动查 PATH 与常见安装位置）")
    parser.add_argument("--settle", type=float, default=0.5,
                        help="每份落地后停顿秒数（默认 0.5）")
    parser.add_argument("--timeout", type=float, default=90.0,
                        help="单文件等待落地超时秒数（默认 90）")
    parser.add_argument("--fix", action="store_true",
                        help="打印前先修复 OneNote「正在清理…」卡死（杀进程+重启，不删数据）")
    parser.add_argument("--write", action="store_true", help="真正打印（默认 dry-run 只预览）")
    parser.add_argument("--list", action="store_true", help="只读探查：打印分区与每页标题后退出")
    args = parser.parse_args()

    if args.fix:
        print("打印前修复 OneNote（杀进程 + 重启 + 等就绪，不删任何文件）：")
        fix_relaunch()

    client = OneNoteClient()

    if args.list:
        cmd_list(client, args.notebook, args.section_group)
        return

    if not args.root or not args.root.is_dir():
        print(f"找不到文件夹根目录：{args.root}（请用 --root 指定 books-done/书名_拆分）")
        return

    # write 模式先定位 SumatraPDF，失败即退出（不打印任何东西）
    sumatra = None
    if args.write:
        try:
            sumatra = find_sumatra(args.sumatra)
        except FileNotFoundError as e:
            sys.exit(f"✗ {e}")

    notebooks = client.get_hierarchy()
    nb = client.find_notebook(notebooks, args.notebook)
    if not nb:
        print(f"未找到笔记本「{args.notebook}」。可先用 --list 查看可用笔记本。")
        return

    scope = resolve_scope(client, nb, args.section_group)
    if scope is None:
        return
    scope_sections, scope_desc = scope

    mode = "打印" if args.write else "dry-run（不打印）"
    print(f"笔记本：{nb.name}　范围：{scope_desc}　模式：{mode}　打印机：{args.printer}")
    if sumatra:
        print(f"打印引擎：{sumatra}")

    sec_by_num = {}
    for sec in scope_sections:
        num = section_number(sec.name, args.section_prefix)
        if num is not None:
            sec_by_num[num] = sec
    if not sec_by_num:
        print(f"未找到以「{args.section_prefix}」开头的分区。")
        return

    total = {"printed": 0, "error": 0}
    for num in sorted(sec_by_num):
        sec = sec_by_num[num]
        folder = args.root / f"{num:02d}"
        if not folder.is_dir():
            _sp(f"\n[{sec.name}] ⚠ 找不到对应文件夹 {folder}，跳过。")
            total["error"] += 1
            continue
        if not args.write:
            pdfs = sorted_pdfs(folder)
            _sp(f"\n[{sec.name}] ⇄ {folder.name}/　{len(pdfs)} 个文件")
            for p in pdfs:
                _sp(f"    · {p.name}")
            continue
        s = print_section(client, sec, folder, printer=args.printer, sumatra=sumatra,
                          settle=args.settle, timeout=args.timeout)
        for k in total:
            total[k] += s[k]

    if not args.write:
        print("\n[dry-run] 未打印。确认映射无误后加 --write 正式打印。")
        return

    print(f"\n══ 总计：打印 {total['printed']}　异常 {total['error']} ══")
    print("收尾：跑 toc-onenote-titles 改标题/删占位页，例如")
    sg = f' --section-group "{args.section_group}"' if args.section_group else ""
    print(f"  uv run toc-onenote-titles{sg} --section-prefix={args.section_prefix} "
          f"--root {args.root} --delete-placeholders --write")


if __name__ == "__main__":
    main()
