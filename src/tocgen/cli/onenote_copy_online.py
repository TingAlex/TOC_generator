"""toc-onenote-copy-online —— Pipeline 5：OneNote 原生整页复制到 OneDrive。

同名本地/在线笔记本按 path 自动区分。目标使用直属同名分区；源可来自笔记本直属
分区，也可用 ``--source-section-group`` 指定本地分区组。默认 dry-run 做完整只读预检。

    toc-onenote-copy-online --source-notebook "书名" --source-section-group "书名"
    toc-onenote-copy-online --source-notebook "书名" --source-section-group "书名" \
        --create-target --write

正式执行时通过 OneNote 自身的“移动或复制页”原生复制整页（不重建 XML），每一页都必须
保留 XPS 打印源并完成同步、等待和复读校验，成功后才处理下一页。
"""

from __future__ import annotations

import argparse
import sys

from ..onenote.client import PI_ALL, Notebook, OneNoteClient, Page, Section
from ..onenote.copy import (
    CopyValidationError,
    PageSignature,
    is_online_path,
    page_signature,
    verify_resume_prefix,
    wait_section_ready,
)
from ..onenote.native_copy import copy_native_page


def _pick_unique(notebooks: list[Notebook], name: str, *, online: bool) -> Notebook:
    matches = [item for item in notebooks
               if item.name == name and is_online_path(item.path) == online]
    kind = "在线" if online else "本地"
    if not matches:
        raise CopyValidationError(f"找不到{kind}笔记本「{name}」。")
    if len(matches) > 1:
        paths = "\n".join(f"  · {item.path}" for item in matches)
        raise CopyValidationError(
            f"找到多个同名{kind}笔记本「{name}」，无法安全判定：\n{paths}")
    return matches[0]


def _pick_optional_online(notebooks: list[Notebook], name: str) -> Notebook | None:
    matches = [item for item in notebooks
               if item.name == name and is_online_path(item.path)]
    if len(matches) > 1:
        paths = "\n".join(f"  · {item.path}" for item in matches)
        raise CopyValidationError(f"找到多个同名在线目标「{name}」：\n{paths}")
    return matches[0] if matches else None


def _pick_reference(notebooks: list[Notebook], name: str | None) -> Notebook:
    online = [item for item in notebooks if is_online_path(item.path)]
    if name:
        matches = [item for item in online if item.name == name]
        if len(matches) != 1:
            raise CopyValidationError(
                f"需要唯一的在线参考笔记本「{name}」，实际找到 {len(matches)} 个。")
        return matches[0]
    if not online:
        raise CopyValidationError("没有在线笔记本可用于确定 OneDrive 父目录。")
    return online[0]


def _unique_sections(sections: list[Section], context: str) -> dict[str, Section]:
    result: dict[str, Section] = {}
    for section in sections:
        if section.name in result:
            raise CopyValidationError(f"{context}存在重名分区「{section.name}」。")
        result[section.name] = section
    return result


def _source_sections(client: OneNoteClient, notebook: Notebook,
                     section_group: str | None) -> list[Section]:
    if section_group:
        group = client.find_section_group(notebook, section_group)
        if group is None:
            raise CopyValidationError(
                f"本地笔记本「{notebook.name}」中找不到分区组「{section_group}」。")
        return group.sections
    if notebook.direct_sections:
        return notebook.direct_sections
    if notebook.section_groups:
        names = "、".join(group.name for group in notebook.section_groups)
        raise CopyValidationError(
            "源笔记本没有直属分区；请用 --source-section-group 指定一个分区组。"
            f"可用：{names}")
    return []


def _scan_sources(client: OneNoteClient, sections: list[Section]) \
        -> dict[str, PageSignature]:
    """写入前遍历全部源页，确认每页都有 XPS 后端；只保留小型摘要。"""
    signatures: dict[str, PageSignature] = {}
    total = sum(len(section.pages) for section in sections)
    done = 0
    for section in sections:
        for page in section.pages:
            done += 1
            source_xml = client.get_page_content(page.id, PI_ALL)
            signatures[page.id] = page_signature(source_xml)
            print(f"  预检源页 {done}/{total}：[{section.name}] {page.name}")
    return signatures


def _without_placeholder(client: OneNoteClient, pages: list[Page]) \
        -> tuple[list[Page], Page | None]:
    if pages and client.is_blank_placeholder(pages[0]):
        return pages[1:], pages[0]
    return pages, None


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(line_buffering=True)
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--source-notebook", required=True, help="源本地笔记本名")
    parser.add_argument("--target-notebook", default=None,
                        help="目标在线笔记本名（默认与源同名）")
    parser.add_argument("--source-section-group", default=None,
                        help="只复制该本地分区组的直属分区；缺省复制笔记本直属分区")
    parser.add_argument("--create-target", action="store_true",
                        help="在线笔记本不存在时创建；也补齐缺失的目标直属分区")
    parser.add_argument("--ref-notebook", default=None,
                        help="创建在线笔记本时，用它确定 OneDrive 同级父目录（缺省取第一个在线本）")
    parser.add_argument("--sync-settle", type=float, default=8.0,
                        help="每页首次同步后的稳定等待秒数（默认 8）")
    parser.add_argument("--ready-timeout", type=float, default=90.0,
                        help="原生复制、新页落地及复读校验超时秒数（默认 90）")
    parser.add_argument("--write", action="store_true",
                        help="真正创建/复制（默认 dry-run，只读完整预检）")
    args = parser.parse_args()

    if args.sync_settle < 0 or args.ready_timeout <= 0:
        sys.exit("✗ --sync-settle 必须 ≥ 0，--ready-timeout 必须 > 0。")

    target_name = args.target_notebook or args.source_notebook
    client = OneNoteClient()
    try:
        notebooks = client.get_hierarchy()
        source_notebook = _pick_unique(
            notebooks, args.source_notebook, online=False)
        source_sections = _source_sections(
            client, source_notebook, args.source_section_group)
        if not source_sections:
            raise CopyValidationError("源作用域内没有分区。")
        source_by_name = _unique_sections(source_sections, "源作用域")

        target_notebook = _pick_optional_online(notebooks, target_name)
        reference: Notebook | None = None
        if target_notebook is None and not args.create_target:
            raise CopyValidationError(
                f"找不到在线笔记本「{target_name}」；确认预览后加 --create-target 创建。")
        if target_notebook is None:
            reference = _pick_reference(notebooks, args.ref_notebook)

        mode = "写入" if args.write else "dry-run（只读）"
        scope = args.source_section_group or "笔记本根"
        print(f"源：本地笔记本「{source_notebook.name}」/ {scope}")
        print(f"目标：在线笔记本「{target_name}」　模式：{mode}")
        print(f"分区：{', '.join(source_by_name)}")
        print(f"总页数：{sum(len(item.pages) for item in source_sections)}\n")

        print("══ 源页 XPS 原生复制预检 ══")
        source_signatures = _scan_sources(client, source_sections)

        # 先只读验证所有已存在目标分区；任何冲突都在创建/删除前暴露。
        target_by_name: dict[str, Section] = {}
        resume: dict[str, int] = {}
        placeholders: dict[str, Page] = {}
        missing: list[str] = []
        if target_notebook is not None:
            target_by_name = _unique_sections(
                target_notebook.direct_sections, "目标笔记本根")
            print("\n══ 目标续跑预检 ══")
            for name, source_section in source_by_name.items():
                destination = target_by_name.get(name)
                if destination is None:
                    missing.append(name)
                    print(f"  [{name}] 缺失，将新建")
                    continue
                pages, placeholder = _without_placeholder(
                    client, client.list_section_pages(destination.id))
                if placeholder is not None:
                    placeholders[name] = placeholder
                resume[name] = verify_resume_prefix(
                    client, source_section, destination, source_signatures,
                    current_pages=pages)
                note = "；将移除开头空白占位页" if placeholder else ""
                print(f"  [{name}] 已有有效页 {resume[name]}/{len(source_section.pages)}{note}")
        else:
            missing = list(source_by_name)
            print(f"\n在线笔记本不存在，将新建并创建 {len(missing)} 个直属分区。")

        if missing and target_notebook is not None and not args.create_target:
            raise CopyValidationError(
                "目标缺少分区且未启用 --create-target：" + ", ".join(missing))

        if not args.write:
            remaining = sum(len(section.pages) - resume.get(name, 0)
                            for name, section in source_by_name.items())
            print(f"\n[dry-run] 预检通过；待线性复制 {remaining} 页，未改动 OneNote。")
            suffix = " --create-target" if target_notebook is None or missing else ""
            print(f"确认后加{suffix} --write 正式执行。")
            return

        if target_notebook is None:
            assert reference is not None
            target_id = client.create_online_notebook(target_name, reference.path)
            target_notebook = Notebook(id=target_id, name=target_name)
            print(f"\n✓ 已创建在线笔记本「{target_name}」（同级于「{reference.name}」）")

        # 所有既有内容均已验证，才执行增量结构变更。
        for name in missing:
            section_id = client.create_section(target_notebook.id, name)
            target_by_name[name] = Section(section_id, name)
            resume[name] = 0
            print(f"✓ 已创建目标直属分区「{name}」")
            initial_pages = wait_section_ready(
                client, section_id, timeout=args.ready_timeout)
            effective_pages, placeholder = _without_placeholder(client, initial_pages)
            if effective_pages:
                raise CopyValidationError(
                    f"刚创建的分区「{name}」意外已有非空页面，已中止且未删除该分区。")
            if placeholder is not None:
                client.delete_page(placeholder.id)
                client.sync_hierarchy(target_notebook.id)
                print(f"  ✓ 新分区空白占位页已移入回收站")

        for name, placeholder in placeholders.items():
            client.delete_page(placeholder.id)
            client.sync_hierarchy(target_notebook.id)
            print(f"✓ [{name}] 开头空白占位页已移入回收站")

        total_remaining = sum(len(section.pages) - resume.get(name, 0)
                              for name, section in source_by_name.items())
        copied = 0
        print(f"\n══ OneNote 原生逐页复制（待复制 {total_remaining} 页）══")
        for name, source_section in source_by_name.items():
            destination = target_by_name[name]
            start = resume.get(name, 0)
            for page_index, source_page in enumerate(
                    source_section.pages[start:], start=start + 1):
                print(f"  [{name}] {page_index}/{len(source_section.pages)} "
                      f"{source_page.name} …", flush=True)
                copy_native_page(
                    client, source_page, destination,
                    target_notebook.id, target_notebook.name,
                    sync_settle=args.sync_settle,
                    ready_timeout=args.ready_timeout,
                    expected_signature=source_signatures[source_page.id],
                )
                copied += 1
                print(f"    ✓ OneNote 原生整页复制，XPS/图片/同步均已复读校验"
                      f"（本次 {copied}/{total_remaining}）")

        # 最终再查一遍完整页序与图片摘要；不以“命令无异常”代替交付校验。
        client.sync_hierarchy(target_notebook.id)
        print("\n══ 最终完整校验 ══")
        for name, source_section in source_by_name.items():
            destination = target_by_name[name]
            count = verify_resume_prefix(
                client, source_section, destination, source_signatures)
            if count != len(source_section.pages):
                raise CopyValidationError(
                    f"分区 {name!r} 最终只有 {count}/{len(source_section.pages)} 页。")
            print(f"  ✓ [{name}] {count} 页，标题/顺序/XPS/图片内容/版面一致")
        print(f"\n══ 完成：本次复制 {copied} 页，在线总页数 "
              f"{sum(len(item.pages) for item in source_sections)} ══")
    except CopyValidationError as exc:
        sys.exit(f"✗ {exc}")


if __name__ == "__main__":
    main()
