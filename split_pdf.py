"""
按 TOC 目录拆分 PDF，并按页数上限分批装入编号子文件夹。

命令行用法（单本调试）：
    python split_pdf.py BOOK_NAME [选项]

批量用法请使用 split_all.py（从 Excel 读取配置）。

示例：
    python split_pdf.py "中学教材全解 高中数学必修第四册" --level 2 --max-pages 100
    python split_pdf.py "某书" --level 1 --max-pages 200 --prefix-digits 2 --prefix-sep _ --folder-digits 3
"""

import argparse
import math
import re
import sys
from pathlib import Path

import fitz  # pymupdf


def _sp(s: str) -> None:
    enc = getattr(sys.stdout, "encoding", "utf-8") or "utf-8"
    try:
        print(s)
    except UnicodeEncodeError:
        print(s.encode(enc, errors="replace").decode(enc))


BASE_DIR = Path(__file__).parent
BOOKS_WORK = BASE_DIR / "books-work"
BOOKS_DONE = BASE_DIR / "books-done"
BOOKS_TODO = BASE_DIR / "books-todo"


def parse_toc(toc_path: Path) -> list[tuple[int, str, int, int]]:
    """Parse toc_parsed.txt → list of (level, title, page, line_no)."""
    entries = []
    with toc_path.open(encoding="utf-8") as f:
        for line_no, raw in enumerate(f, 1):
            line = raw.strip()
            if not line:
                continue
            parts = line.split("|")
            if len(parts) != 3:
                print(f"警告：第 {line_no} 行格式不符（期望 level|title|page），已跳过：{line!r}")
                continue
            try:
                level = int(parts[0])
                title = parts[1].strip()
                page = int(parts[2])
            except ValueError:
                print(f"警告：第 {line_no} 行无法解析数字，已跳过：{line!r}")
                continue
            entries.append((level, title, page, line_no))
    return entries


def sanitize_filename(name: str) -> str:
    """Replace Windows-invalid filename characters with underscore."""
    return re.sub(r'[\\/*?:"<>|]', "_", name)


def run_split(
    book_name: str,
    *,
    level: int = 3,
    max_pages: int = 100,
    max_pages_per_file: int | None = None,
    prefix_digits: int = 3,
    prefix_sep: str = "-",
    folder_digits: int = 2,
    offset: int | None = None,
) -> int:
    """
    拆分一本书的 PDF。

    offset: 印刷页码偏移量（PDF页码 = 印刷页码 + offset）。
            传 None 则自动从 state.json 读取，找不到则为 0。
    返回生成的文件数。
    """
    toc_path = BOOKS_WORK / book_name / "toc_parsed.txt"
    out_root = BOOKS_DONE / f"{book_name}_拆分"

    # 源 PDF：优先 books-done（带书签版），回退到 books-todo（原始版）
    pdf_path = BOOKS_DONE / f"{book_name}.pdf"
    if not pdf_path.exists():
        pdf_path = BOOKS_TODO / f"{book_name}.pdf"

    if not toc_path.exists():
        raise FileNotFoundError(f"找不到 TOC 文件：{toc_path}")
    if not pdf_path.exists():
        raise FileNotFoundError(f"在 books-done/ 和 books-todo/ 中均未找到：{book_name}.pdf")

    # 读取页码偏移量（从 registry / Excel 获取）
    if offset is None:
        import registry as _reg
        _state = _reg.load().get(book_name + ".pdf", {})
        offset = _state.get("offset", 0) or 0

    print(f"页码偏移量：{offset}（印刷页码 + {offset} = PDF页码）")

    # 解析并过滤 TOC（保留 level <= 指定层级的所有条目）
    all_entries = parse_toc(toc_path)
    filtered = [(lv, title, page, line_no)
                for lv, title, page, line_no in all_entries
                if lv <= level]

    if not filtered:
        raise ValueError(f"toc_parsed.txt 中没有第 {level} 层及以上条目")

    # 校验页码：严格递减视为错误
    for i in range(1, len(filtered)):
        prev_lv, prev_title, prev_page, _ = filtered[i - 1]
        lv, title, page, line_no = filtered[i]
        if page < prev_page:
            raise ValueError(
                f"第 {line_no} 行「{title}」(level {lv}) 页码 {page} "
                f"小于上一条「{prev_title}」(level {prev_lv}) 的页码 {prev_page}。"
                f"请先修复 {toc_path}"
            )

    src = fitz.open(str(pdf_path))
    total_pages = len(src)
    print(f"源 PDF：{pdf_path.name}，共 {total_pages} 页")
    print(f"最大拆分层级：{level}，条目数：{len(filtered)}")
    print(f"每文件夹页数上限：{max_pages}\n")

    def save_section(pdf_start: int, pdf_end: int, out_path: Path) -> list[Path]:
        """Save a page range to out_path; splits into _1/_2/… if max_pages_per_file is set."""
        page_count = pdf_end - pdf_start + 1
        if not max_pages_per_file or page_count <= max_pages_per_file:
            out_doc = fitz.open()
            out_doc.insert_pdf(src, from_page=pdf_start - 1, to_page=pdf_end - 1)
            out_doc.save(str(out_path))
            out_doc.close()
            return [out_path]
        # File-level split: stem_1.pdf, stem_2.pdf, …
        n_parts = math.ceil(page_count / max_pages_per_file)
        parts = []
        for i in range(n_parts):
            part_start = pdf_start + i * max_pages_per_file
            part_end   = min(part_start + max_pages_per_file - 1, pdf_end)
            part_path  = out_path.parent / f"{out_path.stem}_{i + 1}{out_path.suffix}"
            out_doc = fitz.open()
            out_doc.insert_pdf(src, from_page=part_start - 1, to_page=part_end - 1)
            out_doc.save(str(part_path))
            out_doc.close()
            parts.append(part_path)
        return parts

    # 计算每个条目的页面范围（印刷页码）
    sections: list[tuple[str, int, int, int]] = []  # (title, pdf_start, pdf_end, level)
    for i, (lv, title, page, _) in enumerate(filtered):
        start = page
        end = total_pages - offset

        # 先看紧挨的下一条（任意层级）：
        #   同页 → 当前条目仅占这一页（章/节标题页与子条目共页的情况）
        #   不同页 → 按层级向后找同级或父级条目，取其起始页 - 1
        if i + 1 < len(filtered) and filtered[i + 1][2] == start:
            end = start
        else:
            for j in range(i + 1, len(filtered)):
                if filtered[j][0] <= lv:
                    end = filtered[j][2] - 1
                    break
        pdf_start = start + offset
        if pdf_start > total_pages:
            print(f"  警告：「{title}」PDF页码 {pdf_start} 超出总页数 {total_pages}，已跳过")
            continue
        pdf_end = min(end + offset, total_pages)
        sections.append((title, pdf_start, pdf_end, lv))

    # 批次文件夹分配与输出
    folder_idx = 1
    cumulative = 0
    total_files = 0

    def next_folder(page_count: int) -> str:
        nonlocal folder_idx, cumulative
        if cumulative + page_count > max_pages and cumulative > 0:
            folder_idx += 1
            cumulative = 0
        return f"{folder_idx:0{folder_digits}d}"

    # 000- 前言（PDF 第1页到第一个 TOC 条目之前）
    if filtered and offset > 0:
        fm_pdf_end = min(filtered[0][2] + offset - 1, total_pages)
        if fm_pdf_end >= 1:
            fm_pages = fm_pdf_end
            folder_name = next_folder(fm_pages)
            out_dir = out_root / folder_name
            out_dir.mkdir(parents=True, exist_ok=True)
            fm_filename = f"000{prefix_sep}{sanitize_filename(book_name)}.pdf"
            saved_fm = save_section(1, fm_pdf_end, out_dir / fm_filename)
            if len(saved_fm) == 1:
                _sp(f"  [{folder_name}] {fm_filename}  ({fm_pages} 页，PDF第1–{fm_pdf_end}页)")
            else:
                _sp(f"  [{folder_name}] {fm_filename}  ({fm_pages} 页→{len(saved_fm)}份，PDF第1–{fm_pdf_end}页)")
                for part_path in saved_fm:
                    _sp(f"    └ {part_path.name}")
            cumulative += fm_pages
            total_files += len(saved_fm)

    # 正文各条目
    for seq, (title, pdf_start, pdf_end, lv) in enumerate(sections, start=1):
        page_count = pdf_end - pdf_start + 1
        folder_name = next_folder(page_count)
        out_dir = out_root / folder_name
        out_dir.mkdir(parents=True, exist_ok=True)

        indent = "  " * (lv - 1)
        prefix = f"{seq:0{prefix_digits}d}{prefix_sep}"
        filename = prefix + sanitize_filename(title) + ".pdf"
        saved = save_section(pdf_start, pdf_end, out_dir / filename)

        if len(saved) == 1:
            _sp(f"  [{folder_name}] {indent}{filename}  ({page_count} 页，PDF第{pdf_start}–{pdf_end}页)")
        else:
            _sp(f"  [{folder_name}] {indent}{filename}  ({page_count} 页→{len(saved)}份，PDF第{pdf_start}–{pdf_end}页)")
            for part_path in saved:
                _sp(f"    └ {part_path.name}")
        cumulative += page_count
        total_files += len(saved)

    src.close()
    print(f"\n完成！共生成 {total_files} 个文件，输出至：{out_root}")
    return total_files


def main() -> None:
    parser = argparse.ArgumentParser(description="按目录拆分 PDF（单本调试用）")
    parser.add_argument("book_name", help="书名（与 books-done/ 和 books-work/ 中的子目录名一致）")
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
                        help="手动指定页码偏移量（不指定则从 state.json 读取）")
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
