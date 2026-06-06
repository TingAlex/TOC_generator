"""
PDF 底层工具：页码规格解析、渲染为图片、写书签、文件名净化。

按章节拆分的「编排」逻辑见 split.py；本模块只放无状态的纯函数。
"""

import re

import fitz  # pymupdf


def parse_page_spec(spec: str) -> list[int]:
    """解析页码规格 '2-4' / '2,4,6' / '2-4,6' → 1-based 页码列表。"""
    pages: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            pages.extend(range(int(start), int(end) + 1))
        elif part:
            pages.append(int(part))
    return pages


def sanitize_filename(name: str) -> str:
    """把 Windows 非法文件名字符替换为下划线。"""
    return re.sub(r'[\\/*?:"<>|]', "_", name)


def render_pages_to_images(pdf_path: str, page_numbers: list[int],
                           dpi: int = 300) -> list[tuple[int, bytes]]:
    """
    渲染指定页（1-based）为 PNG 字节。返回 [(页码, png_bytes), …]。
    超出文档范围的页打印警告并跳过。
    """
    doc = fitz.open(pdf_path)
    total = len(doc)
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    results: list[tuple[int, bytes]] = []
    for pn in page_numbers:
        if pn < 1 or pn > total:
            print(f"  警告：第 {pn} 页超出文档范围（共 {total} 页），已跳过。")
            continue
        pix = doc[pn - 1].get_pixmap(matrix=mat, alpha=False)
        results.append((pn, pix.tobytes("png")))
    doc.close()
    return results


def write_bookmarks(pdf_path: str, toc_entries: list[dict], offset: int,
                    output_path: str) -> int:
    """
    给 PDF 写书签并另存到 output_path。

    toc_entries: [{"level", "title", "page"(印刷页码)}]
    offset: PDF页码 = 印刷页码 + offset
    返回成功写入的书签数（页码越界的条目被跳过）。
    """
    doc = fitz.open(pdf_path)
    total = len(doc)
    toc = []
    skipped = 0
    for entry in toc_entries:
        title = entry.get("title", "").strip()
        if not title:
            continue
        pdf_page = entry.get("page", 0) + offset
        if pdf_page < 1 or pdf_page > total:
            print(f"  警告：「{title}」计算出的 PDF 页码 {pdf_page} 越界（共 {total} 页），已跳过。")
            skipped += 1
            continue
        toc.append([entry.get("level", 1), title, pdf_page])

    doc.set_toc(toc)
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    if skipped:
        print(f"  共写入 {len(toc)} 条书签，{skipped} 条因页码越界被跳过。")
    return len(toc)
