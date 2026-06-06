"""
从 OneNote 指定分区中删除「作为附件误插入的源文件」（默认只删 .pdf），保留打印出来的页面图片。

背景：用 OneNote Batch 导入拆分文件夹时，即便勾选「不插入 PDF 源文件」，源 PDF 仍被作为附件
嵌进了每一页，导致笔记本体积暴涨。本工具逐分区、逐页找出这些附件并删除（仅删附件对象，不动图片）。

本地离线操作（OneNote 桌面 COM 接口），不走网络，符合“关闭同步”状态。删除走 OneNote 内部对象删除，
可在 OneNote 中 Ctrl+Z 撤销或从页面历史/回收站恢复。

用法：
    # 1) 只读探查：列出目标分区每页识别出的附件（名字 + 来源路径），不删
    uv run python onenote_strip_files.py --sections "新分区1" --list

    # 2) dry-run 预览（默认，不写）：列出将删除的附件对照表
    uv run python onenote_strip_files.py --sections "新分区1"

    # 3) 正式删除（建议先在单个分区验证）
    uv run python onenote_strip_files.py --sections "新分区1" --write

    # 4) 批量
    uv run python onenote_strip_files.py --sections "新分区1,新分区2,新分区3" --write

分区按**分区名**精确匹配（逗号分隔）。--ext 可指定扩展名（默认 pdf，逗号分隔）。
"""

import argparse

from onenote_client import OneNoteClient, Section

DEFAULT_NOTEBOOK = "高中数学教辅"


def parse_exts(raw: str) -> set[str]:
    """'pdf,docx' / '.pdf' → {'.pdf', '.docx'}。"""
    out = set()
    for part in raw.split(","):
        p = part.strip().lower()
        if not p:
            continue
        out.add(p if p.startswith(".") else "." + p)
    return out


def process_section(client: OneNoteClient, sec: Section,
                    exts: set[str], list_only: bool, write: bool) -> dict:
    """处理单个分区，返回统计。"""
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
        print(f"    小结：{verb} {stats['deleted'] if write and not list_only else stats['hits']} 个附件")
    return stats


def main():
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--notebook", default=DEFAULT_NOTEBOOK)
    parser.add_argument("--sections", default="",
                        help="目标分区名，逗号分隔（如 '新分区1,新分区2'）")
    parser.add_argument("--ext", default="pdf",
                        help="要删除的附件扩展名，逗号分隔（默认 pdf）")
    parser.add_argument("--list", action="store_true", dest="list_only",
                        help="只读探查：列出每页识别出的附件后退出，不删")
    parser.add_argument("--write", action="store_true",
                        help="真正删除（默认 dry-run 只预览）")
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

    if args.list_only:
        mode = "只读探查（--list）"
    elif args.write:
        mode = "写入（删除）"
    else:
        mode = "dry-run（不写）"
    print(f"笔记本：{nb.name}　模式：{mode}　扩展名：{'/'.join(sorted(exts))}")

    sec_by_name = {sec.name: sec for sec in nb.sections}

    total = {"hits": 0, "deleted": 0, "missing": 0}
    for name in want_names:
        sec = sec_by_name.get(name)
        if sec is None:
            print(f"\n[{name}] ⚠ 笔记本中找不到此分区，跳过。")
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
