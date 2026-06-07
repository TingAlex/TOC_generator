"""
按目录拆分 PDF：依 toc_parsed.txt 把整本切成「章/节/小节」子 PDF，
再以**文件**为单位贪心装进编号子文件夹（01/、02/…），每个文件夹总页数**硬上限** max_pages。

页码转换：PDF页码 = 印刷页码 + offset。

装箱策略（详见 ARCHITECTURE.md）：
  1. 先把每个章节切成输出文件（含前言页；单章超 max_pages_per_file 时切为 _1/_2/… 多份，
     每份 ≤ max_pages_per_file）。
  2. 再按文件顺序贪心装箱：当前文件夹放不下下一个文件就另起新文件夹。
  只要 max_pages_per_file ≤ max_pages（典型 20 ≤ 100），每个文件夹必然 ≤ max_pages——
  优先保证「文件夹（= OneNote 分区 = 同步单元）大小」这一硬约束；代价是一章的多个 _N
  文件可能落到不同文件夹（这对“按大小同步”的场景是可接受且更安全的）。
"""

import math
import sys

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


def _section_ranges(filtered: list[dict], offset: int, total_pages: int) -> list[tuple]:
    """算出每个条目的 PDF 页面范围 → [(title, pdf_start, pdf_end, level)]。"""
    sections = []
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
    return sections


def _plan_files(sections: list[tuple], *, book: str, filtered: list[dict],
                offset: int, total_pages: int, max_pages_per_file: int | None,
                prefix_digits: int, prefix_sep: str) -> list[dict]:
    """把章节展开为有序的输出文件列表（含前言页 + 各章的 _N 切片）。

    每个文件 dict: {filename, start, end, pages, indent}。单文件页数 ≤ max_pages_per_file。
    """
    files: list[dict] = []

    def add_sliced(base_name: str, start: int, end: int, indent: str) -> None:
        count = end - start + 1
        if not max_pages_per_file or count <= max_pages_per_file:
            files.append({"filename": f"{base_name}.pdf", "start": start, "end": end,
                          "pages": count, "indent": indent})
            return
        n_parts = math.ceil(count / max_pages_per_file)
        for i in range(n_parts):
            s = start + i * max_pages_per_file
            e = min(s + max_pages_per_file - 1, end)
            files.append({"filename": f"{base_name}_{i + 1}.pdf", "start": s, "end": e,
                          "pages": e - s + 1, "indent": indent})

    # 000- 前言（PDF 第1页到第一个条目之前），offset>0 时才有
    if filtered and offset > 0:
        fm_end = min(filtered[0]["page"] + offset - 1, total_pages)
        if fm_end >= 1:
            add_sliced(f"000{prefix_sep}{sanitize_filename(book)}", 1, fm_end, "")

    # 正文各条目（序号全局递增）
    for seq, (title, pdf_start, pdf_end, lv) in enumerate(sections, start=1):
        base = f"{seq:0{prefix_digits}d}{prefix_sep}{sanitize_filename(title)}"
        add_sliced(base, pdf_start, pdf_end, "  " * (lv - 1))

    return files


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
    print(f"每文件夹页数上限：{max_pages}（硬上限）"
          + (f"，单文件页数上限：{max_pages_per_file}" if max_pages_per_file else "") + "\n")

    # 1) 章节页面范围 → 2) 展开为输出文件列表（每个文件 ≤ max_pages_per_file）
    sections = _section_ranges(filtered, offset, total_pages)
    planned = _plan_files(sections, book=book, filtered=filtered, offset=offset,
                          total_pages=total_pages, max_pages_per_file=max_pages_per_file,
                          prefix_digits=prefix_digits, prefix_sep=prefix_sep)

    # 3) 以文件为单位贪心装箱（硬上限 max_pages）并写出
    folder_idx, cumulative, total_files = 1, 0, 0
    for f in planned:
        # 放不下且当前文件夹非空 → 另起新文件夹（cumulative>0 守卫：不留空文件夹）
        if cumulative + f["pages"] > max_pages and cumulative > 0:
            folder_idx += 1
            cumulative = 0
        if f["pages"] > max_pages:
            _sp(f"  ⚠ 单文件 {f['filename']} 有 {f['pages']} 页 > 文件夹上限 {max_pages}，"
                f"将独占该文件夹并超限（可调小 --max-pages-per-file）")

        folder_name = f"{folder_idx:0{folder_digits}d}"
        out_dir = out_root / folder_name
        out_dir.mkdir(parents=True, exist_ok=True)

        out_doc = fitz.open()
        out_doc.insert_pdf(src, from_page=f["start"] - 1, to_page=f["end"] - 1)
        out_doc.save(str(out_dir / f["filename"]))
        out_doc.close()

        _sp(f"  [{folder_name}] {f['indent']}{f['filename']}  "
            f"({f['pages']} 页，PDF第{f['start']}–{f['end']}页)")
        cumulative += f["pages"]
        total_files += 1

    src.close()
    print(f"\n完成！共生成 {total_files} 个文件，输出至：{out_root}")
    return total_files
