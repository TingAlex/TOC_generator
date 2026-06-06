"""
Pipeline 1 编排：单本 PDF 的「渲染目录 → OCR → 解析 → 写书签」交互流程。

被 cli.bookmarks（批量）与 cli.bookmarks_one（单本）复用。每完成一步都把 flag
写回 books_config.xlsx，中断重启后自动跳过已完成步骤。
"""

import os
import sys
from pathlib import Path

from . import paths, registry as reg, toc as toc_mod
from .pdf import parse_page_spec, render_pages_to_images, write_bookmarks
from .ai_parse import ocr_pages, parse_toc_text

API_KEYS = ("SILICONFLOW_API_KEY", "DEEPSEEK_API_KEY", "ANTHROPIC_API_KEY", "OPENAI_API_KEY")


def check_api_key() -> None:
    if not any(os.environ.get(k) for k in API_KEYS):
        print("错误：未在 .env 中填入任何 API Key。")
        sys.exit(1)


def print_toc_preview(entries: list[dict]) -> None:
    print(f"\n  识别到 {len(entries)} 条目录项：")
    indent = {1: "", 2: "  ", 3: "    "}
    enc = getattr(sys.stdout, "encoding", "utf-8") or "utf-8"
    for e in entries:
        prefix = indent.get(e["level"], "      ")
        page_str = str(e["page"]) if e["page"] > 0 else "?"
        line = f"  {prefix}[L{e['level']}] {e['title']} ... {page_str}"
        try:
            print(line)
        except UnicodeEncodeError:
            print(line.encode(enc, errors="replace").decode(enc))


def _ask_int(prompt: str) -> int:
    while True:
        try:
            return int(input(prompt).strip())
        except ValueError:
            print("  请输入整数。")


def ask_toc_pages(pdf_name: str) -> list[int] | None:
    while True:
        spec = input(f"\n  《{pdf_name}》目录页范围（如 2-4 或 2,3,4，回车跳过）: ").strip()
        if not spec:
            return None
        try:
            pages = parse_page_spec(spec)
            if pages:
                return pages
        except (ValueError, IndexError):
            pass
        print("  格式有误，请重新输入（示例：2-4 或 2,3,4 或 2-4,6）")


def determine_offset(entries: list[dict]) -> int:
    sample = next((e for e in entries if e["page"] > 0), None)
    if sample is None:
        print("\n  无法找到有效页码，请手动输入偏移量。")
        return _ask_int("  偏移量（PDF页码 - 印刷页码）: ")

    print(f"\n  目录中「{sample['title']}」的印刷页码是第 {sample['page']} 页。")
    while True:
        ans = input("  这一页在 PDF 中实际是第几页？（回车改为手动输入偏移量）: ").strip()
        if not ans:
            return _ask_int("  偏移量（PDF页码 - 印刷页码）: ")
        try:
            pdf_page = int(ans)
            offset = pdf_page - sample["page"]
            print(f"  偏移量 = {pdf_page} - {sample['page']} = {offset}")
            return offset
        except ValueError:
            print("  请输入整数。")


def process_one(pdf_path: Path, registry: dict, *, write: bool = False) -> bool:
    """处理一本书；registry 为全量状态 dict（原地更新并 reg.save）。"""
    book = pdf_path.stem
    key = paths.book_key(book)
    state = registry.setdefault(key, {})
    pages_dir = paths.pages_dir(book)
    pages_dir.parent.mkdir(parents=True, exist_ok=True)
    ocr_raw_path = paths.ocr_raw_path(book)
    toc_parsed_path = paths.toc_parsed_path(book)

    # Step 0：目录页范围
    if not state.get("toc_pages"):
        page_numbers = ask_toc_pages(pdf_path.name)
        if page_numbers is None:
            print("  已跳过。")
            return False
        state["toc_pages"] = page_numbers
        reg.save(registry)
    else:
        page_numbers = state["toc_pages"]
        print(f"\n  目录页范围（已记录）: {page_numbers}")

    # Step 1：渲染目录页
    if not state.get("rendered"):
        print(f"\n  [1/4] 渲染目录页 {page_numbers} ...")
        rendered = render_pages_to_images(str(pdf_path), page_numbers)
        if not rendered:
            print("  错误：没有渲染到任何页面，请检查页码范围。")
            return False
        pages_dir.mkdir(parents=True, exist_ok=True)
        for pn, data in rendered:
            (pages_dir / f"page_{pn:03d}.png").write_bytes(data)
        state["rendered"] = True
        reg.save(registry)
        print(f"  渲染了 {len(rendered)} 页 → {pages_dir}")
    else:
        rendered = [
            (int(p.stem.split("_")[1]), p.read_bytes())
            for p in sorted(pages_dir.glob("page_*.png"))
        ]
        print(f"\n  [1/4] 使用已渲染图片（{len(rendered)} 页）")

    # Step 2：OCR
    if not state.get("ocr_done"):
        print("\n  [2/4] OCR 识别...")
        try:
            ocr_text = ocr_pages(rendered)
        except Exception as e:
            print(f"  错误：OCR 失败 — {e}")
            return False
        ocr_raw_path.write_text(ocr_text, encoding="utf-8")
        state["ocr_done"] = True
        reg.save(registry)
        print(f"  OCR 完成 → {ocr_raw_path}")
    else:
        ocr_text = ocr_raw_path.read_text(encoding="utf-8")
        print("\n  [2/4] 使用已有 OCR 文本")

    # Step 3：解析目录结构
    if not state.get("toc_parsed"):
        print("\n  [3/4] 解析目录结构...")
        try:
            entries = parse_toc_text(ocr_text)
            toc_mod.save(toc_parsed_path, entries)
            state["toc_parsed"] = True
            reg.save(registry)
            print(f"  解析完成 → {toc_parsed_path}")
        except Exception as e:
            print(f"  错误：解析失败 — {e}")
            return False
    else:
        entries = toc_mod.load_file(toc_parsed_path)
        print(f"\n  [3/4] 使用已解析目录（{len(entries)} 条）")

    if not entries:
        print("  错误：目录为空。")
        return False

    print_toc_preview(entries)

    # Step 4：确认偏移量并写书签
    if state.get("bookmarks_added"):
        print(f"\n  [4/4] 书签已添加（{state.get('bookmark_count', '?')} 条），跳过。")
    elif not write:
        if state.get("offset") is None:
            print("\n  [4/4] 确认偏移量（dry-run，不写入 PDF）...")
            state["offset"] = determine_offset(entries)
            reg.save(registry)
        else:
            print(f"\n  [4/4] 偏移量已记录：{state['offset']}（dry-run，不写入 PDF）")
        print("  提示：加 --write 参数可写入 PDF。")
    else:
        if state.get("offset") is not None:
            offset = state["offset"]
            print(f"\n  [4/4] 使用已记录偏移量 {offset}，写入书签...")
        else:
            print("\n  [4/4] 确认偏移量并写入书签...")
            offset = determine_offset(entries)
            state["offset"] = offset
            reg.save(registry)

        output_path = paths.BOOKS_DONE / pdf_path.name
        count = write_bookmarks(str(pdf_path), entries, offset, str(output_path))
        state["bookmarks_added"] = True
        state["bookmark_count"] = count
        reg.save(registry)
        print(f"\n  完成！写入 {count} 条书签 → books-done/{pdf_path.name}")

    return True
