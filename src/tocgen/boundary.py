"""
拆分边界分析：找出「相邻两节」的边界页，渲染其顶部供 Claude 看图判读 fresh / shared。

背景：拆分时第 N 节取 `[起页, 下一节起页-1]`，下一节起页那一**整页**归 PDF2。若下一节在该页
**中间**才开始，则该页顶部是第 N 节的结尾——没进 PDF1 → PDF1 不完整。判读规则：

  · 边界页（= 下一节首页）顶部是「新节标题横幅」→ **fresh**：上一节在上一页就结束了，PDF1 已完整。
  · 顶部是「上一节正文 / 习题残留」→ **shared**：需把这页并入 PDF1（split.py 据 sidecar 把前一节 +1 页）。

本模块只负责「找边界 + 渲染顶部裁剪 + 拼 montage + 打印清单」；判读由 Claude 看图完成，
结果写入 `books-work/{书}/boundary_overlap.txt`（每行 `印刷页|标题`，列出 shared 边界）。
拼图仅用 pymupdf（不引入 Pillow）。
"""

import fitz  # pymupdf

from . import paths

_LABEL_H = 20  # montage 每格顶部标签条高度（px，1pt≈1px 渲染）


def compute_boundaries(filtered: list[dict], offset: int, total_pages: int) -> list[dict]:
    """找出所有需判读的边界条目（= 某节向后找到的「同级/父级」下一条，走 level-scan 分支的）。

    返回按文档顺序、按 `(印刷页, 标题)` 去重的列表：
    `[{page, title, level, pdf_page, prev_title}]`。
    `prev_title` 为文档中紧邻其上的条目标题（= 物理上紧接的 PDF1），仅作清单可读性提示。
    """
    seen: set[tuple[int, str]] = set()
    out: list[dict] = []
    for i, e in enumerate(filtered):
        lv, start = e["level"], e["page"]
        # 「紧邻下一条同页」分支：本条仅占一页、无溢出边界，跳过
        if i + 1 < len(filtered) and filtered[i + 1]["page"] == start:
            continue
        for j in range(i + 1, len(filtered)):
            if filtered[j]["level"] <= lv:
                nb = filtered[j]
                key = (nb["page"], nb["title"])
                pdf_page = nb["page"] + offset
                if key not in seen and 1 <= pdf_page <= total_pages:
                    seen.add(key)
                    out.append({
                        "page": nb["page"], "title": nb["title"], "level": nb["level"],
                        "pdf_page": pdf_page, "prev_title": filtered[j - 1]["title"],
                    })
                break
    return out


def _top_crop(doc: fitz.Document, pdf_page: int, top_frac: float, dpi: int) -> fitz.Pixmap:
    """渲染某页（1-based）顶部 top_frac 的裁剪为 Pixmap。"""
    page = doc[pdf_page - 1]
    r = page.rect
    clip = fitz.Rect(r.x0, r.y0, r.x1, r.y0 + r.height * top_frac)
    return page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), clip=clip, alpha=False)


def _build_montage(cells: list[tuple[fitz.Pixmap, str]], cols: int) -> fitz.Pixmap:
    """把 [(裁剪, 标签)] 拼成一张 montage（grid），返回渲染好的 Pixmap。标签用 ASCII（避免 CJK 字体）。"""
    cw = max(c.width for c, _ in cells)
    ch = max(c.height for c, _ in cells)
    cell_w, cell_h = cw, ch + _LABEL_H
    rows = (len(cells) + cols - 1) // cols
    doc = fitz.open()
    page = doc.new_page(width=cell_w * cols, height=cell_h * rows)
    for k, (pix, label) in enumerate(cells):
        col, row = k % cols, k // cols
        x0, y0 = col * cell_w, row * cell_h
        page.insert_text((x0 + 4, y0 + 14), label, fontsize=11, fontname="helv", color=(0.6, 0, 0))
        img_rect = fitz.Rect(x0, y0 + _LABEL_H, x0 + pix.width, y0 + _LABEL_H + pix.height)
        page.insert_image(img_rect, pixmap=pix)
    out = page.get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False)
    doc.close()
    return out


def render_boundaries(book: str, *, offset: int, level: int = 3,
                      top_frac: float = 0.45, cols: int = 2, per_montage: int = 6,
                      dpi: int = 110) -> list[dict]:
    """渲染该书所有边界页顶部 → books-work/{书}/boundaries/montage_NN.png，并打印清单。

    返回 compute_boundaries 的结果（含每条对应的格子序号），供调用方/人工核对。
    """
    from . import toc as toc_mod  # 局部导入避免循环

    book = paths.stem(book)
    toc_path = paths.toc_parsed_path(book)
    pdf_path = paths.source_pdf(book)
    if not toc_path.exists():
        raise FileNotFoundError(f"找不到 TOC 文件：{toc_path}")

    entries = toc_mod.load_file(toc_path)
    filtered = [e for e in entries if e["level"] <= level]
    if not filtered:
        raise ValueError(f"toc_parsed.txt 中没有第 {level} 层及以上条目")

    doc = fitz.open(str(pdf_path))
    total_pages = len(doc)
    boundaries = compute_boundaries(filtered, offset, total_pages)
    if not boundaries:
        print("没有需判读的边界。")
        doc.close()
        return []

    out_dir = paths.boundaries_dir(book)
    out_dir.mkdir(parents=True, exist_ok=True)
    # 清掉旧 montage，避免与本次清单错位
    for old in out_dir.glob("montage_*.png"):
        old.unlink()

    cells = [(_top_crop(doc, b["pdf_page"], top_frac, dpi),
              f"[{i + 1}] p{b['page']} (PDF {b['pdf_page']})") for i, b in enumerate(boundaries)]
    doc.close()

    print(f"源 PDF：{pdf_path.name}，共 {total_pages} 页；offset={offset}，level={level}")
    print(f"待判读边界：{len(boundaries)} 个（渲染各边界页顶部 {int(top_frac * 100)}%）\n")
    print("序号 | 印刷页 | 下一节（PDF2 首页，待判 fresh/shared） | ← 上一节（PDF1）")
    print("-" * 78)
    for i, b in enumerate(boundaries):
        b["cell"] = i + 1
        print(f"  {i + 1:>2} | {b['page']:>4} | {b['title']}  | {b['prev_title']}")

    n_mon = 0
    for s in range(0, len(cells), per_montage):
        n_mon += 1
        chunk = cells[s:s + per_montage]
        montage = _build_montage(chunk, cols)
        fp = out_dir / f"montage_{n_mon:02d}.png"
        montage.save(str(fp))
        print(f"\n  montage_{n_mon:02d}.png  ← 序号 {s + 1}–{s + len(chunk)}  ({montage.width}x{montage.height})")
    print(f"\n输出至：{out_dir}")
    print("判读后把 shared 边界写入 boundary_overlap.txt（每行 `印刷页|标题`），再重切。")
    return boundaries
