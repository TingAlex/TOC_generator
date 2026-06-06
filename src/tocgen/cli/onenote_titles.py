"""toc-onenote-titles —— Pipeline 4：按拆分文件名核对/改正 OneNote 页标题，
删每个分区开头的空白占位页，并对「误打印两遍」的分区去重。纯本地离线（COM）。

    toc-onenote-titles --list                                   # 只读探查
    toc-onenote-titles --root books-done/书名_拆分               # dry-run 预览
    toc-onenote-titles --root books-done/书名_拆分 --delete-placeholders --dedupe --write
    # 配合 Pipeline 2.5 的「书名分区组」+ 01…0N 分区：
    toc-onenote-titles --section-group "书名" --section-prefix= --root books-done/书名_拆分 --write

分区 ⇄ 文件夹：去前缀后的编号 N ⇄ 文件夹 0N。目标标题 = 去扩展名的完整文件名。
"""

import argparse
from pathlib import Path

from ..onenote.client import OneNoteClient, Section
from ..onenote.common import DEFAULT_NOTEBOOK, section_number, expected_titles, resolve_scope

DEFAULT_SECTION_PREFIX = "新分区"


def cmd_list(client: OneNoteClient, notebook_name: str, section_group: str | None) -> None:
    notebooks = client.get_hierarchy()
    print("可用笔记本：")
    for nb in notebooks:
        mark = " ←" if nb.name == notebook_name else ""
        print(f"  · {nb.name}（{len(nb.sections)} 个分区，{len(nb.section_groups)} 个分区组）{mark}")
    nb = client.find_notebook(notebooks, notebook_name)
    if not nb:
        print(f"\n未找到笔记本「{notebook_name}」。")
        return

    if section_group:
        grp = client.find_section_group(nb, section_group)
        if not grp:
            print(f"\n未找到分区组「{section_group}」。")
            return
        print(f"\n笔记本「{nb.name}」分区组「{grp.name}」分区与页：")
        sections = grp.sections
    else:
        if nb.section_groups:
            print(f"\n笔记本「{nb.name}」分区组：")
            for g in nb.section_groups:
                print(f"  · {g.name}（{len(g.sections)} 个分区）")
        print(f"\n笔记本「{nb.name}」直属/全部分区与页：")
        sections = nb.sections

    for sec in sections:
        print(f"\n  [{sec.name}]（{len(sec.pages)} 页）")
        for i, pg in enumerate(sec.pages, 1):
            print(f"    {i:>3}. {pg.name.strip() or '（空白无标题）'}")


def process_section(client: OneNoteClient, sec: Section, folder: Path,
                    delete_placeholders: bool, dedupe: bool, write: bool) -> dict:
    """核对单个分区，返回统计。打印对照表。"""
    print(f"\n[{sec.name}] ⇄ {folder.name}/")
    stats = {"deleted": 0, "renamed": 0, "ok": 0, "error": 0}
    pages = list(sec.pages)

    # 1) 开头空白无标题占位页
    if pages and client.is_blank_placeholder(pages[0]):
        if write and delete_placeholders:
            client.delete_page(pages[0].id)
            print("    删除占位页：第 1 页（空白无标题）→ 已删（进回收站）")
        else:
            tag = "将删除" if delete_placeholders else "发现（未启用 --delete-placeholders，跳过）"
            print(f"    占位页：第 1 页（空白无标题）→ {tag}")
        if delete_placeholders:
            stats["deleted"] = 1
            pages = pages[1:]

    expected = expected_titles(folder)

    # 2) 去重：页数正好是文件数 2 倍 → 误打印两遍，删后一份
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

    # 3) 页数核对
    if len(pages) != len(expected):
        print(f"    ⚠ 页数({len(pages)}) ≠ 文件数({len(expected)})，本分区中止以防错位对齐。")
        stats["error"] = 1
        return stats

    # 4) 逐页比对标题
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
    parser.add_argument("--delete-placeholders", action="store_true",
                        help="删除每个分区开头的空白无标题占位页")
    parser.add_argument("--dedupe", action="store_true",
                        help="当分区页数正好是文件数 2 倍（误打印两遍）时，删除后一份重复块")
    parser.add_argument("--write", action="store_true", help="真正写入（默认 dry-run 只预览）")
    parser.add_argument("--list", action="store_true", help="只读探查：打印分区组/分区与每页标题后退出")
    args = parser.parse_args()

    client = OneNoteClient()

    if args.list:
        cmd_list(client, args.notebook, args.section_group)
        return

    if not args.root or not args.root.is_dir():
        print(f"找不到文件夹根目录：{args.root}（请用 --root 指定 books-done/书名_拆分）")
        return

    notebooks = client.get_hierarchy()
    nb = client.find_notebook(notebooks, args.notebook)
    if not nb:
        print(f"未找到笔记本「{args.notebook}」。可先用 --list 查看可用笔记本。")
        return

    scope = resolve_scope(client, nb, args.section_group)
    if scope is None:
        return
    scope_sections, scope_desc = scope

    mode = "写入" if args.write else "dry-run（不写）"
    print(f"笔记本：{nb.name}　范围：{scope_desc}　模式：{mode}　"
          f"删占位页：{'是' if args.delete_placeholders else '否'}")

    sec_by_num = {}
    for sec in scope_sections:
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
