"""toc-claude —— Pipeline 1 辅助：渲染目录页 / 写书签，**全程不调用 AI API**。

配合 skill `/toc-by-claude`：中间「看图识目录」由 Claude 自身的多模态能力完成，
直接写出 toc_parsed.txt；本工具只做渲染、写书签两个纯本地步骤。

    toc-claude render "书名" --pages 2-4      # 渲染目录页为 PNG
    toc-claude bookmarks "书名" --offset 18   # 由 toc_parsed.txt 写书签

--pages / --offset 省略时从 books_config.xlsx 已记录值读取。
"""

import argparse
import sys

from .. import paths, registry as reg, toc as toc_mod
from ..pdf import parse_page_spec, render_pages_to_images, write_bookmarks


def _load_state(book: str) -> tuple[dict, dict]:
    registry = reg.load()
    return registry, registry.setdefault(paths.book_key(book), {})


def cmd_render(args: argparse.Namespace) -> None:
    book = paths.stem(args.book)
    try:
        pdf = paths.source_pdf_todo_first(book)
    except FileNotFoundError as e:
        sys.exit(str(e))
    registry, state = _load_state(book)

    pages = parse_page_spec(args.pages) if args.pages else state.get("toc_pages")
    if not pages:
        sys.exit("未指定目录页：请加 --pages（如 2-4 或 2,3,4），"
                 "或先在 books_config.xlsx 的 toc_pages 列填好。")

    rendered = render_pages_to_images(str(pdf), pages, dpi=args.dpi)
    if not rendered:
        sys.exit("没有渲染到任何页面，请检查 --pages 是否在文档范围内。")

    pages_dir = paths.pages_dir(book)
    pages_dir.mkdir(parents=True, exist_ok=True)
    for pn, data in rendered:
        (pages_dir / f"page_{pn:03d}.png").write_bytes(data)

    state["toc_pages"] = pages
    state["rendered"] = True
    reg.save(registry)

    print(f"✓ 渲染 {len(rendered)} 页 → {pages_dir}")
    print(f"  待 Claude 阅读这些 PNG 后写出：{paths.toc_parsed_path(book)}")
    for pn, _ in rendered:
        print(f"    {pages_dir / f'page_{pn:03d}.png'}")


def cmd_bookmarks(args: argparse.Namespace) -> None:
    book = paths.stem(args.book)
    try:
        pdf = paths.source_pdf_todo_first(book)
    except FileNotFoundError as e:
        sys.exit(str(e))
    registry, state = _load_state(book)

    toc_path = paths.toc_parsed_path(book)
    if not toc_path.exists():
        sys.exit(f"找不到 {toc_path}。请先让 Claude 看图写出该文件。")
    entries = toc_mod.load_file(toc_path)
    if not entries:
        sys.exit(f"{toc_path} 为空或格式不对（应为 level|title|页码）。")

    offset = args.offset if args.offset is not None else state.get("offset")
    if offset is None:
        sys.exit("未指定偏移量：请加 --offset N（PDF页码 = 印刷页码 + offset），"
                 "或先在 books_config.xlsx 的 offset 列填好。")

    out = paths.done_pdf(book)
    out.parent.mkdir(parents=True, exist_ok=True)   # 书含系列路径时需建中间层
    count = write_bookmarks(str(pdf), entries, int(offset), str(out))

    state["offset"] = int(offset)
    state["toc_parsed"] = True     # 由 Claude 看图识别完成
    state["bookmarks_added"] = True
    state["bookmark_count"] = count
    reg.save(registry)

    print(f"✓ 写入 {count} 条书签 → {out}")
    print(f"  已更新进度：offset={offset}，toc_parsed/bookmarks_added=True")


def main() -> None:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)

    pr = sub.add_parser("render", help="渲染目录页为 PNG（不调用 API）")
    pr.add_argument("book", help="书名（可带或不带 .pdf）")
    pr.add_argument("--pages", default=None, help="目录页范围，如 2-4 或 2,3,4")
    pr.add_argument("--dpi", type=int, default=300)
    pr.set_defaults(func=cmd_render)

    pb = sub.add_parser("bookmarks", help="由 toc_parsed.txt 写书签（不调用 API）")
    pb.add_argument("book", help="书名（可带或不带 .pdf）")
    pb.add_argument("--offset", type=int, default=None, help="PDF页码 = 印刷页码 + offset")
    pb.set_defaults(func=cmd_bookmarks)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
