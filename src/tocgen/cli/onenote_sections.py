"""toc-onenote-sections —— Pipeline 2.5：按拆分文件夹数预建「分区组 + 空分区」。

读 books-done/{书}_拆分/ 下的 0N 文件夹数，在指定笔记本建一个以书名命名的
分区组 + 同名空分区 01…0N，供 Pipeline 3 打印前准备。纯本地离线（COM）。

书可嵌套在系列文件夹下（books-done/系列/册_拆分/），--book 递归匹配；分区组名取
display_name（路径各段用 `-` 连接，如「薛金星教材全解-人教B-必修第一册」）。

    toc-onenote-sections --book "书名"            # dry-run
    toc-onenote-sections --book "书名" --write     # 正式创建（已有在线笔记本）
    toc-onenote-sections --book "书名" --notebook "新本子" --new-notebook --write
    toc-onenote-sections --book "书名" --local-path "C:\\Users\\用户\\Documents\\笔记本名" --write
"""

import argparse
import sys
from pathlib import Path

from .. import paths
from ..onenote.client import OneNoteClient, Notebook
from ..onenote.common import DEFAULT_NOTEBOOK, section_folder_names


def find_split_folder(root: Path, book: str) -> Path:
    """在 root 下**递归**定位 {书}_拆分 文件夹，支持部分匹配。多于一个则报错列出。

    书可嵌套在系列 / 教辅文件夹下，故匹配的是书 key（相对路径，如
    `薛金星教材全解-人教B/必修第一册`）；`--book 必修第一册` 这类子串照旧能命中。
    """
    if not root.is_dir():
        sys.exit(f"找不到拆分根目录：{root}")
    needle = paths.stem(book)
    candidates = [d for d in paths.iter_split_roots(root)
                  if needle in paths.split_root_key(d, root)]
    if not candidates:
        sys.exit(f"在 {root} 下找不到匹配「{book}」的 *_拆分 文件夹。")
    if len(candidates) > 1:
        names = "\n".join(f"  · {paths.split_root_key(d, root)}" for d in candidates)
        sys.exit(f"「{book}」匹配到多个文件夹，请把 --book 写得更具体：\n{names}")
    return candidates[0]


def pick_ref_online_notebook(notebooks: list[Notebook],
                             ref_name: str | None) -> Notebook | None:
    """选一个在线笔记本（path 以 https:// 开头）作为「同级新建」的参考。"""
    online = [nb for nb in notebooks if nb.path.lower().startswith("https://")]
    if ref_name:
        return next((nb for nb in online if nb.name == ref_name), None)
    return online[0] if online else None


def main() -> None:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--book", required=True,
                        help="书名或书路径，递归定位 books-done/{书}_拆分/（部分匹配）")
    parser.add_argument("--notebook", default=DEFAULT_NOTEBOOK, help="目标笔记本名")
    parser.add_argument("--new-notebook", action="store_true",
                        help="笔记本不存在时新建在线笔记本（同级于现有在线笔记本）")
    parser.add_argument("--local-path", default=None,
                        help="新建本地（不同步）笔记本的磁盘绝对路径（文件夹名即笔记本名）；"
                             "与 --new-notebook 互斥。在线笔记本有 SharePoint 100MB 限制时用此选项。")
    parser.add_argument("--ref-notebook", default=None,
                        help="新建时作「同级参考」的现有在线笔记本名（缺省取第一个在线笔记本）")
    parser.add_argument("--root", type=Path, default=paths.BOOKS_DONE, help="拆分根目录")
    parser.add_argument("--write", action="store_true", help="真正创建（默认 dry-run 只预览）")
    args = parser.parse_args()

    if args.local_path and args.new_notebook:
        sys.exit("✗ --local-path 与 --new-notebook 互斥，请只选其一。")

    folder = find_split_folder(args.root, args.book)
    # 分区组名不能含 `/`，故用 display_name 把路径各段连成一个名字
    book = paths.display_name(paths.split_root_key(folder, args.root))
    secs = section_folder_names(folder)
    if not secs:
        sys.exit(f"文件夹 {folder} 下没有任何 0N 数字子文件夹，无分区可建。")

    nb_display = args.local_path or args.notebook
    mode = "写入" if args.write else "dry-run（不创建）"
    print(f"拆分文件夹：{paths.split_root_key(folder, args.root)}{paths.SPLIT_SUFFIX}")
    print(f"书名（分区组名）：{book}")
    print(f"待建分区（{len(secs)} 个）：{', '.join(secs)}")
    print(f"目标笔记本：{nb_display}　模式：{mode}\n")

    client = OneNoteClient()
    nb_id = None

    if args.local_path:
        # 本地笔记本：直接用磁盘路径创建，不查已有笔记本列表
        print(f"将新建本地笔记本：{args.local_path}")
        if args.write:
            nb_id = client.create_local_notebook(args.local_path)
            print("  → 已新建本地笔记本。")
    else:
        notebooks = client.get_hierarchy()
        nb = client.find_notebook(notebooks, args.notebook)
        if nb is not None:
            nb_id = nb.id
            print(f"✓ 笔记本「{nb.name}」已存在。")
            if any(g.name == book for g in nb.section_groups):  # 只看直属分区组
                sys.exit(f"✗ 笔记本「{nb.name}」下已存在分区组「{book}」，"
                         f"为避免重复已中止（绝不改动已有内容）。")
        else:
            if not args.new_notebook:
                sys.exit(f"✗ 未找到笔记本「{args.notebook}」。"
                         f"加 --new-notebook 可新建在线笔记本，"
                         f"或加 --local-path 新建本地笔记本。")
            ref = pick_ref_online_notebook(notebooks, args.ref_notebook)
            if ref is None:
                if args.ref_notebook:
                    sys.exit(f"✗ 找不到名为「{args.ref_notebook}」的在线笔记本作参考。")
                sys.exit("✗ 没有可作同级参考的在线笔记本（path 以 https:// 开头）。")
            parent = ref.path.rstrip("/").rsplit("/", 1)[0]
            print(f"将新建在线笔记本「{args.notebook}」（同级于「{ref.name}」，父目录 {parent}）")
            if args.write:
                nb_id = client.create_online_notebook(args.notebook, ref.path)
                print("  → 已新建笔记本。")

    if not args.write:
        print("\n[dry-run] 未创建任何东西。确认无误后加 --write 正式创建。")
        return

    sg_id = client.create_section_group(nb_id, book)
    print(f"\n✓ 已建分区组「{book}」")
    for name in secs:
        client.create_section(sg_id, name)
        print(f"  ✓ 分区「{name}」")

    print(f"\n══ 完成：分区组「{book}」+ {len(secs)} 个空分区 ══")
    print("现在可用 toc-onenote-import 把各文件夹的 PDF 打进对应分区，之后用 toc-onenote-titles 收尾。")


if __name__ == "__main__":
    main()
