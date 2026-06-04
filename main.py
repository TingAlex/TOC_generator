#!/usr/bin/env python3
"""
PDF 自动识别目录并批量添加书签

用法：
    将 PDF 文件放入 books-todo/ 目录，然后运行：
    uv run python main.py

中间产物保存在 books-work/{书名}/，进度记录在 registry.json。
手动将 registry.json 中某个字段改为 false 可重做对应步骤。
"""

import argparse
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

import registry as reg
from pdf_utils import parse_page_spec, render_pages_to_images, write_bookmarks
from ai_parser import ocr_pages, parse_toc_text


def check_api_key() -> None:
    keys = ("SILICONFLOW_API_KEY", "DEEPSEEK_API_KEY", "ANTHROPIC_API_KEY", "OPENAI_API_KEY")
    if not any(os.environ.get(k) for k in keys):
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


def _ask_int(prompt: str) -> int:
    while True:
        try:
            return int(input(prompt).strip())
        except ValueError:
            print("  请输入整数。")


def load_entries_from_file(path: Path) -> list[dict]:
    entries = []
    for line in path.read_text(encoding="utf-8").splitlines():
        parts = line.strip().split("|", 2)
        if len(parts) == 3:
            try:
                entries.append({
                    "level": int(parts[0]),
                    "title": parts[1],
                    "page": int(parts[2]),
                })
            except ValueError:
                pass
    return entries


def process_one(pdf_path: Path, done_dir: Path, registry: dict, write: bool = False) -> bool:
    book_name = pdf_path.name
    state = registry.setdefault(book_name, {})
    work_dir = Path("books-work") / pdf_path.stem
    work_dir.mkdir(parents=True, exist_ok=True)
    pages_dir   = work_dir / "pages"
    ocr_raw_path    = work_dir / "ocr_raw.txt"
    toc_parsed_path = work_dir / "toc_parsed.txt"

    # ── Step 0: 目录页范围 ────────────────────────────────────────────
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

    # ── Step 1: 渲染目录页为图片 ──────────────────────────────────────
    if not state.get("rendered"):
        print(f"\n  [1/4] 渲染目录页 {page_numbers} ...")
        rendered = render_pages_to_images(str(pdf_path), page_numbers)
        if not rendered:
            print("  错误：没有渲染到任何页面，请检查页码范围。")
            return False
        pages_dir.mkdir(exist_ok=True)
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

    # ── Step 2: OCR 识别 ──────────────────────────────────────────────
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
        print(f"\n  [2/4] 使用已有 OCR 文本")

    # ── Step 3: 解析目录结构 ──────────────────────────────────────────
    if not state.get("toc_parsed"):
        print("\n  [3/4] 解析目录结构...")
        try:
            entries = parse_toc_text(ocr_text)
            toc_parsed_path.write_text(
                "\n".join(f"{e['level']}|{e['title']}|{e['page']}" for e in entries),
                encoding="utf-8",
            )
            state["toc_parsed"] = True
            reg.save(registry)
            print(f"  解析完成 → {toc_parsed_path}")
        except Exception as e:
            print(f"  错误：解析失败 — {e}")
            return False
    else:
        entries = load_entries_from_file(toc_parsed_path)
        print(f"\n  [3/4] 使用已解析目录（{len(entries)} 条）")

    if not entries:
        print("  错误：目录为空。")
        return False

    print_toc_preview(entries)

    # ── Step 4: 确认偏移量并写入书签 ─────────────────────────────────
    if state.get("bookmarks_added"):
        print(f"\n  [4/4] 书签已添加（{state.get('bookmark_count', '?')} 条），跳过。")
    elif not write:
        # 仍然确认偏移量并保存，但不写 PDF
        if state.get("offset") is None:
            print("\n  [4/4] 确认偏移量（dry-run，不写入 PDF）...")
            offset = determine_offset(entries)
            state["offset"] = offset
            reg.save(registry)
        else:
            offset = state["offset"]
            print(f"\n  [4/4] 偏移量已记录：{offset}（dry-run，不写入 PDF）")
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

        output_path = done_dir / pdf_path.name
        count = write_bookmarks(str(pdf_path), entries, offset, str(output_path))
        state["bookmarks_added"] = True
        state["bookmark_count"] = count
        reg.save(registry)
        print(f"\n  完成！写入 {count} 条书签 → books-done/{pdf_path.name}")

    return True


def main() -> None:
    parser = argparse.ArgumentParser(description="PDF 自动添加书签")
    parser.add_argument("--write", action="store_true",
                        help="真正写入 PDF（默认 dry-run，只跑识别不写文件）")
    args = parser.parse_args()

    todo_dir = Path("books-todo")
    done_dir = Path("books-done")

    check_api_key()

    if not todo_dir.exists():
        print(f"错误：找不到目录 {todo_dir}。")
        sys.exit(1)

    done_dir.mkdir(exist_ok=True)

    pdfs = sorted(todo_dir.glob("*.pdf"))
    if not pdfs:
        print("books-todo/ 中没有 PDF 文件。")
        return

    registry = reg.load()

    print(f"发现 {len(pdfs)} 个待处理文件：")
    for p in pdfs:
        s = registry.get(p.name, {})
        flags = " ".join(
            k for k in ("rendered", "ocr_done", "toc_parsed", "bookmarks_added")
            if s.get(k)
        )
        print(f"  - {p.name}  [{flags or '未开始'}]")

    success = skipped = failed = 0

    for i, pdf_path in enumerate(pdfs, 1):
        print(f"\n{'─' * 55}")
        print(f"  [{i}/{len(pdfs)}] {pdf_path.name}")
        print(f"{'─' * 55}")

        try:
            ok = process_one(pdf_path, done_dir, registry, write=args.write)
            if ok:
                success += 1
            else:
                skipped += 1
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
