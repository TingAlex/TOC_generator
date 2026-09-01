"""toc-boundaries —— 渲染相邻两节的边界页顶部，供 Claude 看图判读 fresh / shared。

    toc-boundaries render "书名" [--level 3] [--offset N] [--top-frac 0.45]

判读后把 shared 边界写入 books-work/{书}/boundary_overlap.txt（每行 `印刷页|标题`），
再用 toc-split / toc-split-all 重切——split 会据此把前一节末尾 +1 页保 PDF1 完整。详见 boundary.py。
"""

import argparse
import sys

from .. import paths, registry
from ..boundary import render_boundaries


def main() -> None:
    parser = argparse.ArgumentParser(description="渲染拆分边界页顶部供判读（fresh/shared）")
    parser.add_argument("command", choices=["render"], help="render：渲染边界页 montage")
    parser.add_argument("book_name",
                        help="书名或书路径，如 薛金星教材全解-人教B/必修第一册；"
                             "唯一命中时可只写片段（必修第一册）")
    parser.add_argument("--level", type=int, choices=[1, 2, 3], default=3,
                        help="拆分最大层级（须与拆分时一致，默认 3）")
    parser.add_argument("--offset", type=int, default=None,
                        help="页码偏移量（不指定则从 books_config.xlsx 读取）")
    parser.add_argument("--top-frac", type=float, default=0.45,
                        help="渲染每个边界页顶部的比例（默认 0.45）")
    parser.add_argument("--cols", type=int, default=2, help="montage 列数（默认 2）")
    parser.add_argument("--per-montage", type=int, default=6, help="每张 montage 的格子数（默认 6）")
    args = parser.parse_args()

    try:
        book = paths.resolve_book(args.book_name)
    except LookupError as e:
        sys.exit(f"错误：{e}")

    offset = args.offset
    if offset is None:
        state = registry.load().get(paths.book_key(book), {})
        offset = state.get("offset", 0) or 0

    try:
        render_boundaries(book, offset=offset, level=args.level,
                          top_frac=args.top_frac, cols=args.cols, per_montage=args.per_montage)
    except (FileNotFoundError, ValueError) as e:
        sys.exit(f"错误：{e}")


if __name__ == "__main__":
    main()
