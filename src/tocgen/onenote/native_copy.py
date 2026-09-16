"""通过 OneNote 自身的“移动或复制页”对话框执行原生整页复制。

桌面 COM 没有 CopyPage 方法；微软为桌面 OneNote 提供的原生入口是 ``Ctrl+Alt+M``。
本模块用 COM 精确导航源页，用 UI Automation 在该原生对话框中选择目标分区和“复制”，
再用 COM 只读校验目标页的 XPS 后端与图片。绝不通过 UpdatePageContent 重建正文。
"""

from __future__ import annotations

from collections.abc import Callable
import time

from .client import PI_ALL, OneNoteClient, Page, Section
from .copy import (
    CopyValidationError,
    PageSignature,
    assert_pages_equal,
    page_signature,
)


def _pywinauto():
    try:
        import warnings
        with warnings.catch_warnings():
            # pywinauto 0.6.9 在 Python 3.14 会发出已知的旧转义序列/COM apartment 提示；
            # 本项目的端到端 OneNote 原生复制测试已验证该组合可用。
            warnings.simplefilter("ignore", SyntaxWarning)
            warnings.filterwarnings("ignore", message="Revert to STA COM threading mode")
            import win32gui
            from pywinauto import Desktop, findwindows, keyboard
    except ImportError as exc:
        raise CopyValidationError(
            "缺少 Windows 原生页面复制依赖 pywinauto；请先运行 uv sync。") from exc
    return Desktop, findwindows, keyboard, win32gui


def _wait_main_window(page_title: str, timeout: float):
    Desktop, _, _, _ = _pywinauto()
    deadline = time.monotonic() + timeout
    expected = f"{page_title} - OneNote"
    while time.monotonic() < deadline:
        matches = [window for window in Desktop(backend="uia").windows()
                   if window.element_info.class_name == "Framework::CFrame"
                   and window.window_text() == expected]
        if len(matches) == 1:
            return matches[0]
        time.sleep(0.25)
    raise CopyValidationError(
        f"无法唯一定位已导航到源页的 OneNote 主窗口：{expected!r}。")


def _wait_copy_dialog(process_id: int, timeout: float):
    Desktop, findwindows, _, _ = _pywinauto()
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        handles = findwindows.find_windows(
            backend="win32", process=process_id, class_name="NUIDialog",
            visible_only=True, enabled_only=True)
        matches = []
        for handle in handles:
            dialog = Desktop(backend="win32").window(handle=handle)
            if dialog.window_text().casefold() in {"移动或复制页", "move or copy pages"}:
                matches.append(handle)
        if len(matches) == 1:
            return matches[0]
        time.sleep(0.25)
    raise CopyValidationError("OneNote 没有打开“移动或复制页”对话框。")


def _direct_tree_children(item) -> list:
    return [child for child in item.children()
            if child.element_info.control_type == "TreeItem"]


def _select_destination(dialog_handle: int, notebook_name: str,
                        section_name: str, timeout: float):
    Desktop, _, _, _ = _pywinauto()
    deadline = time.monotonic() + timeout
    dialog = Desktop(backend="uia").window(handle=dialog_handle)
    candidates = [item for item in dialog.descendants(control_type="TreeItem")
                  if item.window_text() == notebook_name
                  and item.parent().element_info.control_type == "Tree"]
    if not candidates:
        raise CopyValidationError(
            f"OneNote 原生对话框中找不到目标笔记本「{notebook_name}」。")

    for candidate in candidates:
        try:
            candidate.expand()
        except Exception:
            # 已展开或当前 UIA provider 不允许重复展开时，直接检查孩子。
            pass
        time.sleep(0.15)

    while time.monotonic() < deadline:
        dialog = Desktop(backend="uia").window(handle=dialog_handle)
        matches = []
        for candidate in [item for item in dialog.descendants(control_type="TreeItem")
                          if item.window_text() == notebook_name
                          and item.parent().element_info.control_type == "Tree"]:
            direct = [child for child in _direct_tree_children(candidate)
                      if child.window_text() == section_name]
            matches.extend(direct)
        if len(matches) == 1:
            matches[0].select()
            if not matches[0].is_selected():
                raise CopyValidationError("目标分区已找到，但 OneNote 没有选中它。")
            return dialog
        if len(matches) > 1:
            raise CopyValidationError(
                f"同名笔记本中有多个直属分区「{section_name}」，无法安全选择。")
        time.sleep(0.25)
    raise CopyValidationError(
        f"目标笔记本「{notebook_name}」下找不到直属分区「{section_name}」。")


def _cancel_dialog(dialog_handle: int) -> None:
    try:
        Desktop, _, _, _ = _pywinauto()
        dialog = Desktop(backend="uia").window(handle=dialog_handle)
        buttons = [item for item in dialog.descendants(control_type="Button")
                   if item.window_text().casefold() in {"取消", "cancel"}]
        if len(buttons) == 1:
            buttons[0].invoke()
    except Exception:
        pass


def copy_current_page_with_onenote(notebook_name: str, section_name: str,
                                   page_title: str, timeout: float = 30.0) -> None:
    """触发 OneNote 原生“移动或复制页”，选择目标直属分区并点击“复制”。"""
    _, _, keyboard, win32gui = _pywinauto()
    main = _wait_main_window(page_title, timeout)
    if not main.is_enabled():
        raise CopyValidationError("OneNote 主窗口被其它对话框阻塞，请先关闭弹窗。")
    main.set_focus()
    keyboard.send_keys("^%m")
    dialog_handle = _wait_copy_dialog(main.element_info.process_id, timeout)
    try:
        dialog = _select_destination(
            dialog_handle, notebook_name, section_name, timeout)
        buttons = [item for item in dialog.descendants(control_type="Button")
                   if item.window_text().casefold() in {"复制", "copy"}]
        if len(buttons) != 1:
            raise CopyValidationError("OneNote 原生对话框中无法唯一定位“复制”按钮。")
        buttons[0].invoke()
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if not win32gui.IsWindow(dialog_handle):
                return
            time.sleep(0.25)
        raise CopyValidationError("点击“复制”后，OneNote 对话框没有在超时前关闭。")
    except Exception:
        _cancel_dialog(dialog_handle)
        raise


def copy_native_page(
        client: OneNoteClient, source_page: Page, destination_section: Section,
        destination_notebook_id: str, destination_notebook_name: str, *,
        sync_settle: float = 8.0, ready_timeout: float = 90.0,
        expected_signature: PageSignature | None = None,
        copy_action: Callable[[str, str, str, float], None] =
        copy_current_page_with_onenote) -> str:
    """用 OneNote 原生 UI 复制一页，完成同步与 XPS/图片复读后返回新页 ID。"""
    source_xml = client.get_page_content(source_page.id, PI_ALL)
    expected = expected_signature or page_signature(source_xml)
    before = {page.id for page in client.list_section_pages(destination_section.id)}
    new_page: Page | None = None
    try:
        client.navigate_to(source_page.id)
        copy_action(destination_notebook_name, destination_section.name,
                    source_page.name, ready_timeout)

        deadline = time.monotonic() + ready_timeout
        last_error: Exception | None = None
        while time.monotonic() < deadline:
            pages = client.list_section_pages(destination_section.id)
            created = [page for page in pages if page.id not in before]
            if len(created) > 1:
                raise CopyValidationError(
                    f"一次原生复制产生了 {len(created)} 个新页，无法安全判定。")
            if len(created) == 1:
                new_page = created[0]
                try:
                    if new_page.name != source_page.name:
                        raise CopyValidationError(
                            f"标题不一致：{source_page.name!r} != {new_page.name!r}。")
                    actual = page_signature(
                        client.get_page_content(new_page.id, PI_ALL))
                    assert_pages_equal(expected, actual)
                    break
                except Exception as exc:
                    last_error = exc
            time.sleep(0.25)
        else:
            raise CopyValidationError(
                f"原生复制页在 {ready_timeout:g} 秒内未完整落地：{last_error}")

        client.sync_hierarchy(new_page.id)
        client.sync_hierarchy(destination_notebook_id)
        if sync_settle > 0:
            time.sleep(sync_settle)
        client.sync_hierarchy(destination_notebook_id)
        actual = page_signature(client.get_page_content(new_page.id, PI_ALL))
        assert_pages_equal(expected, actual)
        return new_page.id
    except Exception as exc:
        cleanup_error: Exception | None = None
        if new_page is not None:
            try:
                client.delete_page(new_page.id)
                client.sync_hierarchy(destination_notebook_id)
            except Exception as cleanup_exc:
                cleanup_error = cleanup_exc
        suffix = f"；回收本次新页也失败：{cleanup_error}" if cleanup_error else ""
        raise CopyValidationError(
            f"OneNote 原生整页复制 {source_page.name!r} 失败，已中止{suffix}：{exc}") from exc
