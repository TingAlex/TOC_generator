"""toc-split-all —— Pipeline 2 批量：读 Excel，对所有「拆分完成=False」的书执行拆分。

    toc-split-all                # 处理全部待拆分
    toc-split-all --dry-run      # 只列待处理书目
    toc-split-all --book "书名"  # 仅处理书名含此串的书（部分匹配）
"""

import argparse

from .. import paths, bookconfig
from ..bookconfig import read_books_config, read_split_config, to_bool
from ..split import run_split


def main() -> None:
    parser = argparse.ArgumentParser(description="从 Excel 批量拆分 PDF")
    parser.add_argument("--dry-run", action="store_true", help="只列待处理书目，不实际拆分")
    parser.add_argument("--book", type=str, default=None, help="只处理书名含此串的书（部分匹配）")
    args = parser.parse_args()

    split_cfg = read_split_config()
    print(f"全局拆分格式：{split_cfg}\n")

    wb, books, col_idx = read_books_config()
    ws = wb.active

    name_to_row = {}
    for row_no, row in enumerate(ws.iter_rows(min_row=3), start=3):
        if row[0].value:
            name_to_row[str(row[0].value)] = row_no

    pending = [b for b in books if not to_bool(b.get("拆分完成"))]
    if args.book:
        pending = [b for b in pending if args.book in str(b.get("书名", ""))]
    print(f"共 {len(books)} 本书，已完成 {len(books) - len(pending)} 本，待处理 {len(pending)} 本")
    if not pending:
        print("所有书本均已拆分完成。")
        return

    print("\n待处理书目：")
    for b in pending:
        print(f"  · {b['书名']}  (level={int(b.get('split_level') or 3)}, offset={int(b.get('offset') or 0)})")

    if args.dry_run:
        print("\n[dry-run] 未执行拆分。")
        return

    print()
    done_col = col_idx.get("拆分完成")
    succeeded = failed = 0
    for book in pending:
        name = str(book["书名"])
        offset_val = book.get("offset")
        print(f"{'=' * 60}\n正在拆分：{name}\n{'=' * 60}")
        try:
            run_split(
                name,
                level=int(book.get("split_level") or 3),
                max_pages=split_cfg["max_pages"],
                max_pages_per_file=split_cfg["max_pages_per_file"],
                prefix_digits=split_cfg["prefix_digits"],
                prefix_sep=split_cfg["prefix_sep"],
                folder_digits=split_cfg["folder_digits"],
                offset=int(offset_val) if offset_val is not None else None,
            )
            if done_col and name in name_to_row:
                ws.cell(row=name_to_row[name], column=done_col, value=True)
            wb.save(paths.BOOKS_CONFIG_PATH)
            succeeded += 1
            print()
        except Exception as e:
            print(f"\n  [跳过] 拆分失败：{e}\n")
            failed += 1

    print(f"\n批量拆分完成：成功 {succeeded} 本，失败 {failed} 本")


if __name__ == "__main__":
    main()
