"""toc-onenote-strip —— Pipeline 4 子工具：从指定分区删除「误插入的源文件附件」
（默认仅 .pdf），保留打印页图片。纯本地离线（COM），默认 dry-run。

    toc-onenote-strip --sections "新分区 1" --list
    toc-onenote-strip --sections "新分区 1,新分区 2" --write
    toc-onenote-strip --section-group "书名" --sections "01,02" --write
"""

import argparse

from ..onenote.client import OneNoteClient, Section
from ..onenote.common import DEFAULT_NOTEBOOK, resolve_scope


def parse_exts(raw: str) -> set[str]:
    """'pdf,docx' / '.pdf' → {'.pdf', '.docx'}。"""
    out = set()
    for part in raw.split(","):
        p = part.strip().lower()
        if p:
            out.add(p if p.startswith(".") else "." + p)
    return out


def process_section(client: OneNoteClient, sec: Section,
                    exts: set[str], list_only: bool, write: bool) -> dict:
    print(f"\n[{sec.name}]（{len(sec.pages)} 页）")
    stats = {"hits": 0, "deleted": 0}
    for i, pg in enumerate(sec.pages, 1):
        files = client.list_inserted_files(pg.id, exts)
        if not files:
            continue
        stats["hits"] += len(files)
        title = pg.name.strip() or "（空白无标题）"
        for f in files:
            src = f" ← {f.path_source}" if f.path_source else ""
            if list_only:
                print(f"    第{i:>3}页[{title}]：{f.name}{src}")
            elif write:
                client.delete_page_content(pg.id, f.object_id)
                stats["deleted"] += 1
                print(f"    第{i:>3}页[{title}]：{f.name} → 已删")
            else:
                print(f"    第{i:>3}页[{title}]：{f.name} → 待删")

    if stats["hits"] == 0:
        print("    （未发现匹配的附件）")
    else:
        verb = "已删" if write and not list_only else "命中"
        n = stats["deleted"] if write and not list_only else stats["hits"]
        print(f"    小结：{verb} {n} 个附件")
    return stats


def main() -> None:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--notebook", default=DEFAULT_NOTEBOOK)
    parser.add_argument("--sections", default="",
                        help="目标分区名，逗号分隔（如 '新分区 1,新分区 2'）")
    parser.add_argument("--section-group", default=None,
                        help="只在该分区组内按名匹配分区（缺省全部分区）")
    parser.add_argument("--ext", default="pdf",
                        help="要删除的附件扩展名，逗号分隔（默认 pdf）")
    parser.add_argument("--list", action="store_true", dest="list_only",
                        help="只读探查：列出每页识别出的附件后退出，不删")
    parser.add_argument("--write", action="store_true", help="真正删除（默认 dry-run 只预览）")
    args = parser.parse_args()

    exts = parse_exts(args.ext)
    if not exts:
        print("--ext 为空，无可删除的扩展名。")
        return
    want_names = [s.strip() for s in args.sections.split(",") if s.strip()]
    if not want_names:
        print("请用 --sections 指定要处理的分区名（逗号分隔）。")
        return

    client = OneNoteClient()
    notebooks = client.get_hierarchy()
    nb = client.find_notebook(notebooks, args.notebook)
    if not nb:
        print(f"未找到笔记本「{args.notebook}」。可用笔记本：")
        for n in notebooks:
            print(f"  · {n.name}")
        return

    mode = "只读探查（--list）" if args.list_only else ("写入（删除）" if args.write else "dry-run（不写）")
    scope = resolve_scope(client, nb, args.section_group)
    if scope is None:
        return
    scope_sections, scope_desc = scope
    print(f"笔记本：{nb.name}　范围：{scope_desc}　模式：{mode}　扩展名：{'/'.join(sorted(exts))}")

    sec_by_name = {sec.name: sec for sec in scope_sections}
    total = {"hits": 0, "deleted": 0, "missing": 0}
    for name in want_names:
        sec = sec_by_name.get(name)
        if sec is None:
            print(f"\n[{name}] ⚠ 找不到此分区，跳过。")
            total["missing"] += 1
            continue
        s = process_section(client, sec, exts, args.list_only, args.write)
        total["hits"] += s["hits"]
        total["deleted"] += s["deleted"]

    print(f"\n══ 总计：命中 {total['hits']} 个附件，已删 {total['deleted']} 个，"
          f"缺失分区 {total['missing']} 个 ══")
    if not args.list_only and not args.write:
        print("（dry-run，未写入。确认无误后加 --write 正式删除。）")


if __name__ == "__main__":
    main()
