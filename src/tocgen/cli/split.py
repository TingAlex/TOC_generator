"""toc-split —— Pipeline 2 单本（命令行调试）：按 toc_parsed.txt 拆分一本书。

    toc-split "书名" --level 3 --max-pages 100
批量请用 toc-split-all（从 Excel 读配置）。
"""

import argparse
import sys

from ..split import run_split


def main() -> None:
    parser = argparse.ArgumentParser(description="按目录拆分 PDF（单本调试用）")
    parser.add_argument("book_name", help="书名（与 books-work/ 子目录名一致）")
    parser.add_argument("--level", type=int, choices=[1, 2, 3], default=3,
                        help="拆分最大层级（默认 3，即全部）")
    parser.add_argument("--max-pages", type=int, default=100,
                        help="每个批次文件夹页数上限（默认 100）")
    parser.add_argument("--max-pages-per-file", type=int, default=None,
                        help="单个输出文件最大页数，超出则切为 _1/_2/… 多份（默认不限）")
    parser.add_argument("--prefix-digits", type=int, default=3,
                        help="文件前缀序号位数（默认 3 → 001-）")
    parser.add_argument("--prefix-sep", type=str, default="-",
                        help="前缀分隔符（默认 -）")
    parser.add_argument("--folder-digits", type=int, default=2,
                        help="批次文件夹编号位数（默认 2 → 01）")
    parser.add_argument("--offset", type=int, default=None,
                        help="手动指定页码偏移量（不指定则从 Excel 配置读取）")
    args = parser.parse_args()

    try:
        run_split(
            args.book_name,
            level=args.level,
            max_pages=args.max_pages,
            max_pages_per_file=args.max_pages_per_file,
            prefix_digits=args.prefix_digits,
            prefix_sep=args.prefix_sep,
            folder_digits=args.folder_digits,
            offset=args.offset,
        )
    except (FileNotFoundError, ValueError) as e:
        sys.exit(f"错误：{e}")


if __name__ == "__main__":
    main()
