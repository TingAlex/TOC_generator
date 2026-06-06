"""toc-bookmarks-one —— Pipeline 1 单本：只处理 books-todo/ 里的第一本，便于调试。

    toc-bookmarks-one            # dry-run
    toc-bookmarks-one --write    # 写入 PDF
"""

import argparse

from dotenv import load_dotenv

from .. import paths, registry as reg
from ..pipeline1 import process_one, check_api_key


def main() -> None:
    load_dotenv()
    parser = argparse.ArgumentParser(description="Pipeline 1 单本测试")
    parser.add_argument("--write", action="store_true", help="写入 PDF（默认 dry-run）")
    args = parser.parse_args()

    check_api_key()
    paths.BOOKS_DONE.mkdir(exist_ok=True)

    pdfs = sorted(paths.BOOKS_TODO.glob("*.pdf"))
    if not pdfs:
        print("books-todo/ 中没有 PDF 文件。")
        return

    first = pdfs[0]
    print(f"测试文件：{first.name}")
    print(f"模式：{'写入 PDF' if args.write else 'dry-run（不写 PDF）'}\n")
    process_one(first, reg.load(), write=args.write)


if __name__ == "__main__":
    main()
