"""
按本地 PDF 文件名核对/改正 OneNote「高中数学教辅」新分区 1~10 的页标题，
并删除每个分区开头自动生成的空白无标题占位页。

本地离线操作（OneNote 桌面 COM 接口），不走网络，符合“关闭同步”状态。

用法：
    # 1) 只读探查：打印笔记本下全部分区与每页标题
    uv run python onenote_sync_titles.py --list

    # 2) dry-run 预览（默认，不写）：逐分区列出待删占位页 + 标题差异
    uv run python onenote_sync_titles.py

    # 3) 正式执行：删占位页 + 改标题
    uv run python onenote_sync_titles.py --delete-placeholders --write

标题格式：保留完整文件名（去扩展名，含 NNN- 编号前缀与 _N 拆分后缀）。
分区 ⇄ 文件夹映射：新分区N ⇄ 文件夹 0N。
"""

import argparse
import re
from pathlib import Path

from onenote_client import OneNoteClient, Section

DEFAULT_NOTEBOOK = "高中数学教辅"
DEFAULT_SECTION_PREFIX = "新分区"
DEFAULT_ROOT = Path("books-done") / "更高更妙的高中数学思想与方法(第16版)_拆分"

# 文件名前缀编号，如 "117-4.4.1 弦长问题" → 117
_NUM_PREFIX = re.compile(r"^(\d+)")


def section_number(name: str, prefix: str) -> int | None:
    """从分区名解析编号，如 '新分区7' → 7。"""
    if not name.startswith(prefix):
        return None
    rest = name[len(prefix):].strip()
    return int(rest) if rest.isdigit() else None


def expected_titles(folder: Path) -> list[str]:
    """文件夹内 PDF 的期望页标题 = 去扩展名的文件名，按编号前缀数字排序。"""
    def sort_key(p: Path):
        m = _NUM_PREFIX.match(p.stem)
        return (int(m.group(1)) if m else 1 << 30, p.stem)
    return [p.stem for p in sorted(folder.glob("*.pdf"), key=sort_key)]


def cmd_list(client: OneNoteClient, notebook_name: str) -> None:
    notebooks = client.get_hierarchy()
    print("可用笔记本：")
    for nb in notebooks:
        mark = " ←" if nb.name == notebook_name else ""
        print(f"  · {nb.name}（{len(nb.sections)} 个分区）{mark}")
    nb = client.find_notebook(notebooks, notebook_name)
    if not nb:
        print(f"\n未找到笔记本「{notebook_name}」。")
        return
    print(f"\n笔记本「{nb.name}」分区与页：")
    for sec in nb.sections:
        print(f"\n  [{sec.name}]（{len(sec.pages)} 页）")
        for i, pg in enumerate(sec.pages, 1):
            title = pg.name.strip() or "（空白无标题）"
            print(f"    {i:>3}. {title}")


def process_section(client: OneNoteClient, sec: Section, folder: Path,
                    delete_placeholders: bool, dedupe: bool, write: bool) -> dict:
    """核对单个分区，返回统计。打印对照表。"""
    print(f"\n[{sec.name}] ⇄ {folder.name}/")
    stats = {"deleted": 0, "renamed": 0, "ok": 0, "error": 0}

    pages = list(sec.pages)

    # 1) 处理开头的空白无标题占位页
    if pages and client.is_blank_placeholder(pages[0]):
        if write and delete_placeholders:
            client.delete_page(pages[0].id)
            print("    删除占位页：第 1 页（空白无标题）→ 已删（进回收站）")
        else:
            tag = "将删除" if delete_placeholders else "发现（未启用 --delete-placeholders，跳过）"
            print(f"    占位页：第 1 页（空白无标题）→ {tag}")
        if delete_placeholders:
            stats["deleted"] = 1
            pages = pages[1:]  # 删后按剩余页对齐

    expected = expected_titles(folder)

    # 2) 去重：页数正好是文件数 2 倍 → 误打印两遍，删后一份重复块，保留前一份
    if dedupe and expected and len(pages) == 2 * len(expected):
        dup = pages[len(expected):]
        if write:
            for pg in dup:
                client.delete_page(pg.id)
            print(f"    去重：删除后一份重复块 第{len(expected)+1}~{len(pages)}页（{len(dup)}页）→ 已删（进回收站）")
        else:
            print(f"    去重：将删除后一份重复块 第{len(expected)+1}~{len(pages)}页（{len(dup)}页）")
        stats["deleted"] += len(dup)
        pages = pages[:len(expected)]

    # 3) 页数与文件数核对
    if len(pages) != len(expected):
        print(f"    ⚠ 页数({len(pages)}) ≠ 文件数({len(expected)})，本分区中止以防错位对齐。")
        stats["error"] = 1
        return stats

    # 3) 逐页比对标题
    for i, (pg, want) in enumerate(zip(pages, expected), 1):
        cur = pg.name.strip()
        if cur == want:
            stats["ok"] += 1
            continue
        cur_show = cur or "（空白）"
        if write:
            client.set_page_title(pg.id, want)
            print(f"    第{i:>3}页：{cur_show}  →  {want}  [已改]")
        else:
            print(f"    第{i:>3}页：{cur_show}  →  {want}  [待改]")
        stats["renamed"] += 1

    print(f"    小结：√{stats['ok']} 改{stats['renamed']} 删{stats['deleted']}")
    return stats


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--notebook", default=DEFAULT_NOTEBOOK)
    parser.add_argument("--root", type=Path, default=DEFAULT_ROOT)
    parser.add_argument("--section-prefix", default=DEFAULT_SECTION_PREFIX)
    parser.add_argument("--delete-placeholders", action="store_true",
                        help="删除每个分区开头的空白无标题占位页")
    parser.add_argument("--dedupe", action="store_true",
                        help="当分区页数正好是文件数 2 倍（误打印两遍）时，删除后一份重复块")
    parser.add_argument("--write", action="store_true",
                        help="真正写入（默认 dry-run 只预览）")
    parser.add_argument("--list", action="store_true",
                        help="只读探查：打印分区与每页标题后退出")
    args = parser.parse_args()

    client = OneNoteClient()

    if args.list:
        cmd_list(client, args.notebook)
        return

    if not args.root.is_dir():
        print(f"找不到文件夹根目录：{args.root}")
        return

    notebooks = client.get_hierarchy()
    nb = client.find_notebook(notebooks, args.notebook)
    if not nb:
        print(f"未找到笔记本「{args.notebook}」。可先用 --list 查看可用笔记本。")
        return

    mode = "写入" if args.write else "dry-run（不写）"
    print(f"笔记本：{nb.name}　模式：{mode}　"
          f"删占位页：{'是' if args.delete_placeholders else '否'}")

    # 分区编号 → Section
    sec_by_num = {}
    for sec in nb.sections:
        num = section_number(sec.name, args.section_prefix)
        if num is not None:
            sec_by_num[num] = sec

    if not sec_by_num:
        print(f"未找到以「{args.section_prefix}」开头的分区。")
        return

    total = {"deleted": 0, "renamed": 0, "ok": 0, "error": 0}
    for num in sorted(sec_by_num):
        folder = args.root / f"{num:02d}"
        if not folder.is_dir():
            print(f"\n[{sec_by_num[num].name}] ⚠ 找不到对应文件夹 {folder}，跳过。")
            total["error"] += 1
            continue
        s = process_section(client, sec_by_num[num], folder,
                            args.delete_placeholders, args.dedupe, args.write)
        for k in total:
            total[k] += s[k]

    print(f"\n══ 总计：√{total['ok']} 改{total['renamed']} 删{total['deleted']} "
          f"异常分区{total['error']} ══")
    if not args.write:
        print("（dry-run，未写入。确认无误后加 --delete-placeholders --write 正式执行。）")


if __name__ == "__main__":
    main()
