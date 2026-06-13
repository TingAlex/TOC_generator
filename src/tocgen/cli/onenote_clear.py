"""toc-onenote-clear —— 清空分区组内所有分区的全部页面（进 OneNote 回收站，可恢复）。

在重新打印（toc-onenote-import）前用于清除旧页，防止重复内容。纯本地离线（COM）。

    toc-onenote-clear --section-group "书名"                           # dry-run（只预览）
    toc-onenote-clear --section-group "书名" --write                   # 真正删除
    toc-onenote-clear --notebook "笔记本名" --section-group "书名" --write
"""

import argparse
import sys

from ..onenote.client import OneNoteClient
from ..onenote.common import DEFAULT_NOTEBOOK, resolve_scope


def _sp(s: str) -> None:
    enc = getattr(sys.stdout, "encoding", "utf-8") or "utf-8"
    try:
        print(s)
    except UnicodeEncodeError:
        print(s.encode(enc, errors="replace").decode(enc))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--notebook", default=DEFAULT_NOTEBOOK, help="目标笔记本名")
    parser.add_argument("--section-group", default=None,
                        help="只清空该分区组内的分区（强烈建议填写，避免误删其他书的分区）")
    parser.add_argument("--write", action="store_true",
                        help="真正删除（默认 dry-run 只预览）；删除进回收站，可恢复")
    args = parser.parse_args()

    client = OneNoteClient()
    notebooks = client.get_hierarchy()
    nb = client.find_notebook(notebooks, args.notebook)
    if not nb:
        print(f"未找到笔记本「{args.notebook}」。可用：{', '.join(n.name for n in notebooks)}")
        return

    scope = resolve_scope(client, nb, args.section_group)
    if scope is None:
        return
    sections, desc = scope

    mode = "删除（进回收站）" if args.write else "dry-run（不删）"
    _sp(f"笔记本：{nb.name}　范围：{desc}　模式：{mode}")

    total = 0
    for sec in sections:
        pages = client.list_section_pages(sec.id)
        _sp(f"\n  [{sec.name}]  {len(pages)} 页")
        for pg in pages:
            title = pg.name.strip() or "（无标题）"
            if args.write:
                client.delete_page(pg.id)
                _sp(f"    [OK] {title}")
            else:
                _sp(f"    [-] {title}")
            total += 1

    verb = "已删入回收站" if args.write else "（dry-run，未删）"
    print(f"\n共 {total} 页 {verb}")
    if not args.write:
        print("确认无误后加 --write 正式删除。")


if __name__ == "__main__":
    main()
