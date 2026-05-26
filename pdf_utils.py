import fitz  # pymupdf


def parse_page_spec(spec: str) -> list[int]:
    """Parse page specification like '2-4' or '2,4,6' or '2-4,6' into 1-based page numbers."""
    pages = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            pages.extend(range(int(start), int(end) + 1))
        else:
            pages.append(int(part))
    return pages


def render_pages_to_images(pdf_path: str, page_numbers: list[int], dpi: int = 300) -> list[tuple[int, bytes]]:
    """
    Render specified pages (1-based) to PNG bytes.
    Returns list of (page_number, png_bytes) tuples.
    """
    doc = fitz.open(pdf_path)
    total = len(doc)
    results = []
    scale = dpi / 72
    mat = fitz.Matrix(scale, scale)

    for pn in page_numbers:
        if pn < 1 or pn > total:
            print(f"  警告：第 {pn} 页超出文档范围（共 {total} 页），已跳过。")
            continue
        page = doc[pn - 1]
        pix = page.get_pixmap(matrix=mat, alpha=False)
        results.append((pn, pix.tobytes("png")))

    doc.close()
    return results



def write_bookmarks(pdf_path: str, toc_entries: list[dict], offset: int, output_path: str) -> int:
    """
    Write bookmarks to PDF and save to output_path.

    toc_entries: list of {"level": int, "title": str, "page": int}
                 where page is the printed page number from the book.
    offset: PDF_page_number (1-based) - printed_page_number
            so PDF_page = printed_page + offset

    Returns number of bookmarks successfully written.
    """
    doc = fitz.open(pdf_path)
    total = len(doc)
    toc = []
    skipped = 0

    for entry in toc_entries:
        printed = entry.get("page", 0)
        title = entry.get("title", "").strip()
        level = entry.get("level", 1)

        if not title:
            continue

        pdf_page = printed + offset
        if pdf_page < 1:
            print(f"  警告：「{title}」计算出的PDF页码 {pdf_page} 无效，已跳过。")
            skipped += 1
            continue
        if pdf_page > total:
            print(f"  警告：「{title}」计算出的PDF页码 {pdf_page} 超出总页数 {total}，已跳过。")
            skipped += 1
            continue

        toc.append([level, title, pdf_page])

    doc.set_toc(toc)
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()

    written = len(toc)
    if skipped:
        print(f"  共写入 {written} 条书签，{skipped} 条因页码异常被跳过。")
    return written
