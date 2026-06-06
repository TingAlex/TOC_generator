"""
按目录拆分 PDF：依 toc_parsed.txt 把整本切成「章/节/小节」子 PDF，
按页数上限分批装入编号子文件夹（01/、02/…）。

页码转换：PDF页码 = 印刷页码 + offset。
批次分配 / 文件级切片算法见 ARCHITECTURE.md。
"""

import math
import sys
from pathlib import Path

import fitz  # pymupdf

from . import paths, toc as toc_mod, registry
from .pdf import sanitize_filename


def _sp(s: str) -> None:
    """容错打印（控制台编码不支持的字符不致崩溃）。"""
    enc = getattr(sys.stdout, "encoding", "utf-8") or "utf-8"
    try:
        print(s)
    except UnicodeEncodeError:
        print(s.encode(enc, errors="replace").decode(enc))


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
    拆分一本书。offset 传 None 时从 books_config.xlsx 读取，找不到则 0。
    返回生成的文件数。
    """
    book = paths.stem(book_name)
    toc_path = paths.toc_parsed_path(book)
    out_root = paths.split_root(book)
    pdf_path = paths.source_pdf(book)  # 优先 books-done（带书签版）

    if not toc_path.exists():
        raise FileNotFoundError(f"找不到 TOC 文件：{toc_path}")

    if offset is None:
        state = registry.load().get(paths.book_key(book), {})
        offset = state.get("offset", 0) or 0
    print(f"页码偏移量：{offset}（印刷页码 + {offset} = PDF页码）")

    # 解析 + 过滤层级 + 校验页码不递减
    entries = toc_mod.load_file(toc_path)
    filtered = [e for e in entries if e["level"] <= level]
    if not filtered:
        raise ValueError(f"toc_parsed.txt 中没有第 {level} 层及以上条目")
    toc_mod.check_nondecreasing(filtered)

    src = fitz.open(str(pdf_path))
    total_pages = len(src)
    print(f"源 PDF：{pdf_path.name}，共 {total_pages} 页")
    print(f"最大拆分层级：{level}，条目数：{len(filtered)}")
    print(f"每文件夹页数上限：{max_pages}\n")

    def save_section(pdf_start: int, pdf_end: int, out_path: Path) -> list[Path]:
        """保存一个页面区间；若超过 max_pages_per_file 则切为 _1/_2/… 多份。"""
        page_count = pdf_end - pdf_start + 1
        if not max_pages_per_file or page_count <= max_pages_per_file:
            out_doc = fitz.open()
            out_doc.insert_pdf(src, from_page=pdf_start - 1, to_page=pdf_end - 1)
            out_doc.save(str(out_path))
            out_doc.close()
            return [out_path]
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

    # 计算每个条目的 PDF 页面范围
    sections: list[tuple[str, int, int, int]] = []  # (title, pdf_start, pdf_end, level)
    for i, e in enumerate(filtered):
        lv, title, start = e["level"], e["title"], e["page"]
        end = total_pages - offset
        # 紧邻下一条同页 → 本条仅占一页；否则向后找同级/父级条目的起始页 - 1
        if i + 1 < len(filtered) and filtered[i + 1]["page"] == start:
            end = start
        else:
            for j in range(i + 1, len(filtered)):
                if filtered[j]["level"] <= lv:
                    end = filtered[j]["page"] - 1
                    break
        pdf_start = start + offset
        if pdf_start > total_pages:
            print(f"  警告：「{title}」PDF页码 {pdf_start} 超出总页数 {total_pages}，已跳过")
            continue
        pdf_end = min(end + offset, total_pages)
        sections.append((title, pdf_start, pdf_end, lv))

    # 批次文件夹分配
    folder_idx = 1
    cumulative = 0
    total_files = 0

    def next_folder(page_count: int) -> str:
        nonlocal folder_idx, cumulative
        if cumulative + page_count > max_pages and cumulative > 0:
            folder_idx += 1
            cumulative = 0
        return f"{folder_idx:0{folder_digits}d}"

    # 000- 前言（PDF 第1页到第一个条目之前），offset>0 时才有
    if filtered and offset > 0:
        fm_pdf_end = min(filtered[0]["page"] + offset - 1, total_pages)
        if fm_pdf_end >= 1:
            folder_name = next_folder(fm_pdf_end)
            out_dir = out_root / folder_name
            out_dir.mkdir(parents=True, exist_ok=True)
            fm_filename = f"000{prefix_sep}{sanitize_filename(book)}.pdf"
            saved_fm = save_section(1, fm_pdf_end, out_dir / fm_filename)
            if len(saved_fm) == 1:
                _sp(f"  [{folder_name}] {fm_filename}  ({fm_pdf_end} 页，PDF第1–{fm_pdf_end}页)")
            else:
                _sp(f"  [{folder_name}] {fm_filename}  ({fm_pdf_end} 页→{len(saved_fm)}份，PDF第1–{fm_pdf_end}页)")
                for part_path in saved_fm:
                    _sp(f"    └ {part_path.name}")
            cumulative += fm_pdf_end
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
