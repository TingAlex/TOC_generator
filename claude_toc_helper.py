"""
Claude-OCR 流水线辅助：渲染目录页 / 写书签 —— **全程不调用任何 AI API**。

配合 skill `toc-by-claude` 使用：用 Claude 对话自身的多模态能力替代项目里的
DeepSeek-OCR + DeepSeek-V3（即 ai_parser.py / llm_client.py 那条 API 链路）。
本脚本只做「渲染」和「写书签」两个纯本地步骤，中间的「看图识目录」由 Claude 完成。

子命令：
    # 1) 渲染目录页为 PNG（供 Claude 阅读）
    uv run python claude_toc_helper.py render "书名" --pages 2-4
        → books-work/{书名}/pages/page_NNN.png ；记录 toc_pages + rendered

    # 2) 由 toc_parsed.txt 写书签（Claude 写好 toc_parsed.txt 之后）
    uv run python claude_toc_helper.py bookmarks "书名" --offset 18
        → books-done/{书名}.pdf ；记录 offset / ocr_done / toc_parsed /
          bookmarks_added / bookmark_count

说明：
- 书名可带或不带 `.pdf`。源 PDF 优先取 books-todo/，回退 books-done/。
- --pages / --offset 省略时，从 books_config.xlsx 已记录值读取。
- toc_parsed.txt 行格式：`{level}|{title}|{印刷页码}`（与项目其余部分一致）。
"""

import argparse
import sys
from pathlib import Path

import registry as reg
from pdf_utils import parse_page_spec, render_pages_to_images, write_bookmarks

BOOKS_TODO = Path("books-todo")
BOOKS_DONE = Path("books-done")
BOOKS_WORK = Path("books-work")


def _stem(book: str) -> str:
    return book[:-4] if book.lower().endswith(".pdf") else book


def _source_pdf(stem: str) -> Path:
    """源 PDF：优先 books-todo/（原始版），回退 books-done/（已带书签版）。"""
    for d in (BOOKS_TODO, BOOKS_DONE):
        p = d / f"{stem}.pdf"
        if p.exists():
            return p
    sys.exit(f"找不到 PDF：books-todo/ 或 books-done/ 下均无 {stem}.pdf")


def _load_state(book_key: str) -> tuple[dict, dict]:
    """返回 (整个 registry, 该书 state)。book_key 形如 '书名.pdf'。"""
    registry = reg.load()
    return registry, registry.setdefault(book_key, {})


def cmd_render(args: argparse.Namespace) -> None:
    stem = _stem(args.book)
    book_key = f"{stem}.pdf"
    pdf = _source_pdf(stem)
    registry, state = _load_state(book_key)

    if args.pages:
        pages = parse_page_spec(args.pages)
    else:
        pages = state.get("toc_pages")
    if not pages:
        sys.exit("未指定目录页：请加 --pages（如 2-4 或 2,3,4），"
                 "或先在 books_config.xlsx 的 toc_pages 列填好。")

    rendered = render_pages_to_images(str(pdf), pages, dpi=args.dpi)
    if not rendered:
        sys.exit("没有渲染到任何页面，请检查 --pages 是否在文档范围内。")

    pages_dir = BOOKS_WORK / stem / "pages"
    pages_dir.mkdir(parents=True, exist_ok=True)
    for pn, data in rendered:
        (pages_dir / f"page_{pn:03d}.png").write_bytes(data)

    state["toc_pages"] = pages
    state["rendered"] = True
    reg.save(registry)

    print(f"✓ 渲染 {len(rendered)} 页 → {pages_dir}")
    print("  待 Claude 阅读这些 PNG 后写出："
          f"{BOOKS_WORK / stem / 'toc_parsed.txt'}")
    for pn, _ in rendered:
        print(f"    {pages_dir / f'page_{pn:03d}.png'}")


def _load_entries(toc_path: Path) -> list[dict]:
    entries = []
    for ln in toc_path.read_text(encoding="utf-8").splitlines():
        parts = ln.strip().split("|", 2)
        if len(parts) == 3:
            try:
                entries.append({"level": int(parts[0]),
                                "title": parts[1],
                                "page": int(parts[2])})
            except ValueError:
                pass
    return entries


def cmd_bookmarks(args: argparse.Namespace) -> None:
    stem = _stem(args.book)
    book_key = f"{stem}.pdf"
    pdf = _source_pdf(stem)
    registry, state = _load_state(book_key)

    toc_path = BOOKS_WORK / stem / "toc_parsed.txt"
    if not toc_path.exists():
        sys.exit(f"找不到 {toc_path}。请先让 Claude 看图写出该文件。")
    entries = _load_entries(toc_path)
    if not entries:
        sys.exit(f"{toc_path} 为空或格式不对（应为 level|title|页码）。")

    offset = args.offset if args.offset is not None else state.get("offset")
    if offset is None:
        sys.exit("未指定偏移量：请加 --offset N（PDF页码 = 印刷页码 + offset），"
                 "或先在 books_config.xlsx 的 offset 列填好。")

    BOOKS_DONE.mkdir(exist_ok=True)
    out = BOOKS_DONE / f"{stem}.pdf"
    count = write_bookmarks(str(pdf), entries, int(offset), str(out))

    state["offset"] = int(offset)
    state["ocr_done"] = True       # 由 Claude 完成（替代 DeepSeek-OCR）
    state["toc_parsed"] = True     # 由 Claude 完成（替代 DeepSeek-V3）
    state["bookmarks_added"] = True
    state["bookmark_count"] = count
    reg.save(registry)

    print(f"✓ 写入 {count} 条书签 → {out}")
    print(f"  已更新进度：offset={offset}，ocr_done/toc_parsed/bookmarks_added=True")


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
    pb.add_argument("--offset", type=int, default=None,
                    help="PDF页码 = 印刷页码 + offset")
    pb.set_defaults(func=cmd_bookmarks)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
