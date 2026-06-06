"""toc-bookmarks —— Pipeline 1 批量：识别 books-todo/ 中所有 PDF 的目录并加书签。

    toc-bookmarks            # dry-run，只跑识别不写 PDF
    toc-bookmarks --write    # 正式写入书签到 books-done/
"""

import argparse
import sys

from dotenv import load_dotenv

from .. import paths, registry as reg
from ..pipeline1 import process_one, check_api_key


def main() -> None:
    load_dotenv()
    parser = argparse.ArgumentParser(description="PDF 自动添加书签（批量）")
    parser.add_argument("--write", action="store_true",
                        help="真正写入 PDF（默认 dry-run，只跑识别不写文件）")
    args = parser.parse_args()

    check_api_key()

    if not paths.BOOKS_TODO.exists():
        sys.exit(f"错误：找不到目录 {paths.BOOKS_TODO}。")
    paths.BOOKS_DONE.mkdir(exist_ok=True)

    pdfs = sorted(paths.BOOKS_TODO.glob("*.pdf"))
    if not pdfs:
        print("books-todo/ 中没有 PDF 文件。")
        return

    registry = reg.load()
    print(f"发现 {len(pdfs)} 个待处理文件：")
    for p in pdfs:
        s = registry.get(paths.book_key(p.stem), {})
        flags = " ".join(k for k in ("rendered", "ocr_done", "toc_parsed", "bookmarks_added")
                         if s.get(k))
        print(f"  - {p.name}  [{flags or '未开始'}]")

    success = skipped = failed = 0
    for i, pdf_path in enumerate(pdfs, 1):
        print(f"\n{'─' * 55}\n  [{i}/{len(pdfs)}] {pdf_path.name}\n{'─' * 55}")
        try:
            ok = process_one(pdf_path, registry, write=args.write)
            success, skipped = (success + 1, skipped) if ok else (success, skipped + 1)
        except KeyboardInterrupt:
            print("\n\n  已中断。")
            break
        except Exception as e:
            print(f"  未预期的错误：{e}")
            failed += 1

    print(f"\n{'═' * 55}")
    print(f"  全部完成：成功 {success} 本，跳过 {skipped} 本，失败 {failed} 本。")


if __name__ == "__main__":
    main()
