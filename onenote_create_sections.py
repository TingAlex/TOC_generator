"""
Pipeline 2.5：按拆分文件夹数，在 OneNote 预建「分区组 + 空分区」。

读 `books-done/{书名}_拆分/` 下的 `0N` 子文件夹数量，在指定笔记本里建一个
**以书名命名的分区组**，并在组内建好对应数量的空分区（命名与文件夹同名：01…0N）。
之后用户只需用 OneNote Batch 把每个文件夹插入对应分区即可。

本地离线操作（OneNote 桌面 COM 接口），不走网络 Graph，符合“关闭同步”状态。
（注：`--new-notebook` 新建的是**在线**笔记本——做成现有在线笔记本的同级，改动随同步上传。）

用法：
    $env:PYTHONUTF8=1   # 让中文输出不乱码

    # 1) dry-run 预览（默认，不创建任何东西）
    uv run python onenote_create_sections.py --book "书名"

    # 2) 确认无误后正式创建
    uv run python onenote_create_sections.py --book "书名" --write

    # 3) 目标笔记本不存在 → 新建在线笔记本（同级于现有在线笔记本）
    uv run python onenote_create_sections.py --book "书名" --notebook "新本子" --new-notebook --write

安全：默认 dry-run；创建为纯增量（只新增空分区组/空分区，不删不改既有内容）；
若目标笔记本已存在同名（=书名）分区组，则中止并报警，绝不改动已有内容。
"""

import argparse
import re
import sys
from pathlib import Path

from onenote_client import OneNoteClient, Notebook

DEFAULT_NOTEBOOK = "高中数学教辅"
DEFAULT_ROOT = Path("books-done")
SPLIT_SUFFIX = "_拆分"

# 纯数字文件夹名，如 "01"、"02"
_NUM_DIR = re.compile(r"^\d+$")


def find_split_folder(root: Path, book: str) -> Path:
    """在 root 下定位 {书名}_拆分 文件夹，支持部分匹配。多于一个则报错列出。"""
    if not root.is_dir():
        sys.exit(f"找不到拆分根目录：{root}")

    candidates = [d for d in root.iterdir()
                  if d.is_dir() and d.name.endswith(SPLIT_SUFFIX) and book in d.name]
    if not candidates:
        sys.exit(f"在 {root} 下找不到匹配「{book}」的 *_拆分 文件夹。")
    if len(candidates) > 1:
        names = "\n".join(f"  · {d.name}" for d in candidates)
        sys.exit(f"「{book}」匹配到多个文件夹，请把 --book 写得更具体：\n{names}")
    return candidates[0]


def book_name_of(folder: Path) -> str:
    """从 {书名}_拆分 文件夹名还原书名（作为分区组名）。"""
    return folder.name[:-len(SPLIT_SUFFIX)]


def section_names(folder: Path) -> list[str]:
    """收集拆分文件夹下的 0N 子文件夹名（即分区名），按数字排序。"""
    dirs = [d.name for d in folder.iterdir() if d.is_dir() and _NUM_DIR.match(d.name)]
    return sorted(dirs, key=lambda n: int(n))


def pick_ref_online_notebook(notebooks: list[Notebook],
                             ref_name: str | None) -> Notebook | None:
    """选一个在线笔记本（path 以 https:// 开头）作为「同级新建」的参考。"""
    online = [nb for nb in notebooks if nb.path.lower().startswith("https://")]
    if ref_name:
        for nb in online:
            if nb.name == ref_name:
                return nb
        return None
    return online[0] if online else None


def main() -> None:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--book", required=True,
                        help="书名，定位 books-done/{书名}_拆分/（支持部分匹配）")
    parser.add_argument("--notebook", default=DEFAULT_NOTEBOOK, help="目标笔记本名")
    parser.add_argument("--new-notebook", action="store_true",
                        help="笔记本不存在时新建在线笔记本（同级于现有在线笔记本）")
    parser.add_argument("--ref-notebook", default=None,
                        help="新建时作「同级参考」的现有在线笔记本名（缺省自动取第一个在线笔记本）")
    parser.add_argument("--root", type=Path, default=DEFAULT_ROOT, help="拆分根目录")
    parser.add_argument("--write", action="store_true",
                        help="真正创建（默认 dry-run 只预览）")
    args = parser.parse_args()

    # 1) 定位文件夹 + 推导分区名
    folder = find_split_folder(args.root, args.book)
    book = book_name_of(folder)
    secs = section_names(folder)
    if not secs:
        sys.exit(f"文件夹 {folder} 下没有任何 0N 数字子文件夹，无分区可建。")

    mode = "写入" if args.write else "dry-run（不创建）"
    print(f"拆分文件夹：{folder.name}")
    print(f"书名（分区组名）：{book}")
    print(f"待建分区（{len(secs)} 个）：{', '.join(secs)}")
    print(f"目标笔记本：{args.notebook}　模式：{mode}\n")

    client = OneNoteClient()
    notebooks = client.get_hierarchy()
    nb = client.find_notebook(notebooks, args.notebook)

    # 2) 解析/新建笔记本
    nb_id = None
    if nb is not None:
        nb_id = nb.id
        print(f"✓ 笔记本「{nb.name}」已存在。")
        # 3) 分区组查重（仅已存在的笔记本可能撞名；只看直属分区组）
        if any(g.name == book for g in nb.section_groups):
            sys.exit(f"✗ 笔记本「{nb.name}」下已存在分区组「{book}」，"
                     f"为避免重复已中止（绝不改动已有内容）。")
    else:
        if not args.new_notebook:
            sys.exit(f"✗ 未找到笔记本「{args.notebook}」。"
                     f"加 --new-notebook 可新建（在线，同级于现有笔记本）。")
        ref = pick_ref_online_notebook(notebooks, args.ref_notebook)
        if ref is None:
            if args.ref_notebook:
                sys.exit(f"✗ 找不到名为「{args.ref_notebook}」的在线笔记本作参考。")
            sys.exit("✗ 没有可作同级参考的在线笔记本（path 以 https:// 开头）。")
        parent = ref.path.rstrip("/").rsplit("/", 1)[0]
        print(f"将新建在线笔记本「{args.notebook}」（同级于「{ref.name}」，父目录 {parent}）")
        if args.write:
            nb_id = client.create_online_notebook(args.notebook, ref.path)
            print(f"  → 已新建笔记本。")

    # 4) 计划已在上面打印；dry-run 到此为止
    if not args.write:
        print("\n[dry-run] 未创建任何东西。确认无误后加 --write 正式创建。")
        return

    # 5) 创建分区组 + 各分区
    sg_id = client.create_section_group(nb_id, book)
    print(f"\n✓ 已建分区组「{book}」")
    for name in secs:
        client.create_section(sg_id, name)
        print(f"  ✓ 分区「{name}」")

    print(f"\n══ 完成：分区组「{book}」+ {len(secs)} 个空分区 ══")
    print("现在可用 OneNote Batch 把各文件夹插入对应分区，"
          "之后用 onenote_sync_titles.py 收尾。")


if __name__ == "__main__":
    main()
