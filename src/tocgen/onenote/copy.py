"""把本地 OneNote 打印页逐页复制到在线笔记本。

这里只复制可自包含的 ``PageSettings``、``Title`` 与 ``Image`` XML。OneNote
打印页还带有仅在源笔记本有效的 ``XPSFile``、``CallbackID``、objectID 等引用；
跨笔记本复用完整 XML 会触发 ``0x80042014 (hrObjectDoesNotExist)``，因此必须剥离。

每页严格执行：创建目标页 → 写内容 → 同步页/笔记本 → 等待 → 再同步 →
复读并核验全部图片。任一步失败都只删除本次新建页（进回收站），然后中止。
"""

from __future__ import annotations

import base64
import copy
from dataclasses import dataclass
import hashlib
import math
import time
import xml.etree.ElementTree as ET

from .client import ONE_NS, PI_ALL, OneNoteClient, Page, Section

ET.register_namespace("one", ONE_NS)

_COPY_CHILDREN = {"PageSettings", "Title", "Image"}
_DROP_CHILDREN = {"XPSFile", "CallbackID", "QuickStyleDef"}
_DROP_ATTRIBUTES = {
    "objectID",
    "lastModifiedTime",
    "creationTime",
    "lastModifiedBy",
    "lastModifiedByInitials",
    "author",
    "authorInitials",
    "xpsFileIndex",
    "originalPageNumber",
    "isPrintOut",
}


class CopyValidationError(RuntimeError):
    """源页不适合无损复制，或目标页复读校验失败。"""


@dataclass(frozen=True)
class ImageSignature:
    """不保存大块图片本体的可比较摘要。"""

    sha256: str
    format: str
    x: float
    y: float
    width: float
    height: float


def is_online_path(path: str) -> bool:
    """OneNote 层级中的 HTTP(S) path 表示 OneDrive/SharePoint 笔记本。"""
    return path.lower().startswith(("https://", "http://"))


def _local_name(name: str) -> str:
    return name.rsplit("}", 1)[-1]


def _required_float(element: ET.Element | None, attribute: str,
                    context: str) -> float:
    if element is None or element.get(attribute) is None:
        raise CopyValidationError(f"图片缺少 {context}.{attribute}，无法校验无损复制。")
    try:
        return float(element.get(attribute, ""))
    except ValueError as exc:
        raise CopyValidationError(
            f"图片的 {context}.{attribute} 不是数字：{element.get(attribute)!r}") from exc


def image_signatures(page_xml: str) -> tuple[ImageSignature, ...]:
    """提取页内全部图片的内容哈希、格式、位置和尺寸（保持 XML 顺序）。"""
    root = ET.fromstring(page_xml)
    signatures: list[ImageSignature] = []
    for image in root.iter(f"{{{ONE_NS}}}Image"):
        data = image.find(f"{{{ONE_NS}}}Data")
        position = image.find(f"{{{ONE_NS}}}Position")
        size = image.find(f"{{{ONE_NS}}}Size")
        encoded = "" if data is None else "".join((data.text or "").split())
        if not encoded:
            raise CopyValidationError("图片没有内嵌 Data；可能只拿到了本地 XPS 引用。")
        try:
            raw = base64.b64decode(encoded, validate=True)
        except ValueError as exc:
            raise CopyValidationError("图片 Data 不是有效的 Base64。") from exc
        signatures.append(ImageSignature(
            sha256=hashlib.sha256(raw).hexdigest(),
            format=image.get("format", "").lower(),
            x=_required_float(position, "x", "Position"),
            y=_required_float(position, "y", "Position"),
            width=_required_float(size, "width", "Size"),
            height=_required_float(size, "height", "Size"),
        ))
    if not signatures:
        raise CopyValidationError("源页没有图片；本命令只支持 PDF 打印输出页。")
    return tuple(signatures)


def assert_images_equal(expected: tuple[ImageSignature, ...],
                        actual: tuple[ImageSignature, ...],
                        tolerance: float = 0.05) -> None:
    """核验图片二进制与版面；浮点坐标容忍 OneNote 的轻微格式归一化。"""
    if len(expected) != len(actual):
        raise CopyValidationError(
            f"图片数不一致：源页 {len(expected)}，目标页 {len(actual)}。")
    for index, (left, right) in enumerate(zip(expected, actual), 1):
        if left.sha256 != right.sha256 or left.format != right.format:
            raise CopyValidationError(f"第 {index} 张图片的二进制内容或格式不一致。")
        for field in ("x", "y", "width", "height"):
            if not math.isclose(getattr(left, field), getattr(right, field),
                                rel_tol=0.0, abs_tol=tolerance):
                raise CopyValidationError(
                    f"第 {index} 张图片的 {field} 不一致："
                    f"{getattr(left, field)} != {getattr(right, field)}。")


def prepare_printout_page_xml(source_xml: str, destination_page_id: str) -> str:
    """生成可跨笔记本提交的自包含 Page XML，并拒绝可能丢失的其它正文对象。"""
    source = ET.fromstring(source_xml)
    direct_names = [_local_name(child.tag) for child in source]
    unsupported = sorted({name for name in direct_names
                          if name not in _COPY_CHILDREN | _DROP_CHILDREN})
    if unsupported:
        raise CopyValidationError(
            "源页含本命令不会复制的对象：" + ", ".join(unsupported) +
            "。为避免静默丢内容，已中止。")
    if "Title" not in direct_names:
        raise CopyValidationError("源页没有 Title，无法保持页标题。")

    # 同时确认图片确实以内嵌 Data 存在；不能只靠本地 XPSFile 引用。
    image_signatures(source_xml)

    target = ET.Element(f"{{{ONE_NS}}}Page", {"ID": destination_page_id})
    for child in source:
        if _local_name(child.tag) not in _COPY_CHILDREN:
            continue
        cloned = copy.deepcopy(child)
        for node in cloned.iter():
            for attribute in list(node.attrib):
                local_attribute = _local_name(attribute)
                if local_attribute in _DROP_ATTRIBUTES:
                    del node.attrib[attribute]
                elif _local_name(node.tag) == "OE" and local_attribute == "quickStyleIndex":
                    del node.attrib[attribute]
        target.append(cloned)

    body = ET.tostring(target, encoding="unicode", short_empty_elements=True)
    return '<?xml version="1.0"?>' + body


def wait_page_ready(client: OneNoteClient, section_id: str, page_id: str,
                    timeout: float = 30.0, poll: float = 0.25) -> None:
    """等待 CreateNewPage 的页同时出现在层级中且可被 GetPageContent 读取。"""
    deadline = time.monotonic() + timeout
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            if any(page.id == page_id for page in client.list_section_pages(section_id)):
                client.get_page_content(page_id)
                return
        except Exception as exc:  # COM 新对象短暂不可见，保留最后错误供诊断
            last_error = exc
        time.sleep(poll)
    detail = f"；最后错误：{last_error}" if last_error else ""
    raise TimeoutError(f"等待新页 {page_id} 就绪超时（{timeout:g} 秒）{detail}")


def wait_section_ready(client: OneNoteClient, section_id: str,
                       timeout: float = 30.0, poll: float = 0.25) -> list[Page]:
    """等待 OpenHierarchy 新建的分区可以读取，并返回它当时的页列表。"""
    deadline = time.monotonic() + timeout
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            return client.list_section_pages(section_id)
        except Exception as exc:
            last_error = exc
        time.sleep(poll)
    raise TimeoutError(
        f"等待新分区 {section_id} 就绪超时（{timeout:g} 秒）：{last_error}") from last_error


def _wait_verified(client: OneNoteClient, source_page: Page,
                   destination_section: Section, destination_page_id: str,
                   expected: tuple[ImageSignature, ...], timeout: float,
                   poll: float = 0.25) -> None:
    deadline = time.monotonic() + timeout
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            destination_xml = client.get_page_content(destination_page_id, PI_ALL)
            assert_images_equal(expected, image_signatures(destination_xml))
            page = next((item for item in client.list_section_pages(destination_section.id)
                         if item.id == destination_page_id), None)
            if page is None:
                raise CopyValidationError("目标页写入后未出现在目标分区层级中。")
            if page.name != source_page.name:
                raise CopyValidationError(
                    f"标题不一致：源页 {source_page.name!r}，目标页 {page.name!r}。")
            return
        except Exception as exc:
            last_error = exc
        time.sleep(poll)
    raise CopyValidationError(
        f"目标页复读校验在 {timeout:g} 秒内未通过：{last_error}") from last_error


def copy_printout_page(client: OneNoteClient, source_page: Page,
                       destination_section: Section, destination_notebook_id: str,
                       *, sync_settle: float = 8.0,
                       ready_timeout: float = 30.0,
                       expected_signatures: tuple[ImageSignature, ...] | None = None) -> str:
    """线性复制并同步一页；成功返回新页 ID，失败回收本次新页并抛错。"""
    source_xml = client.get_page_content(source_page.id, PI_ALL)
    expected = expected_signatures or image_signatures(source_xml)
    destination_page_id: str | None = None
    try:
        destination_page_id = client.create_new_page(destination_section.id)
        wait_page_ready(client, destination_section.id, destination_page_id,
                        timeout=ready_timeout)
        page_xml = prepare_printout_page_xml(source_xml, destination_page_id)
        client.update_page_content(page_xml)

        # SyncHierarchy 是同步调用；额外稳定等待后再同步/复读，才允许进入下一页。
        client.sync_hierarchy(destination_page_id)
        client.sync_hierarchy(destination_notebook_id)
        if sync_settle > 0:
            time.sleep(sync_settle)
        client.sync_hierarchy(destination_notebook_id)
        _wait_verified(client, source_page, destination_section, destination_page_id,
                       expected, timeout=ready_timeout)
        return destination_page_id
    except Exception as exc:
        cleanup_error: Exception | None = None
        if destination_page_id is not None:
            try:
                client.delete_page(destination_page_id)
                client.sync_hierarchy(destination_notebook_id)
            except Exception as cleanup_exc:
                cleanup_error = cleanup_exc
        suffix = f"；回收本次新页也失败：{cleanup_error}" if cleanup_error else ""
        raise CopyValidationError(
            f"复制页 {source_page.name!r} 失败，已中止{suffix}：{exc}") from exc


def verify_resume_prefix(client: OneNoteClient, source: Section, destination: Section,
                         source_signatures: dict[str, tuple[ImageSignature, ...]],
                         *, current_pages: list[Page] | None = None) -> int:
    """目标必须是源页的完整内容前缀；返回可从哪个下标续跑。"""
    current = (client.list_section_pages(destination.id)
               if current_pages is None else current_pages)
    if len(current) > len(source.pages):
        raise CopyValidationError(
            f"分区 {destination.name!r} 目标页数 {len(current)} 超过源页数 {len(source.pages)}。")
    for index, (left, right) in enumerate(zip(source.pages, current), 1):
        if left.name != right.name:
            raise CopyValidationError(
                f"分区 {source.name!r} 第 {index} 页标题不构成续跑前缀："
                f"{left.name!r} != {right.name!r}。")
        actual = image_signatures(client.get_page_content(right.id, PI_ALL))
        assert_images_equal(source_signatures[left.id], actual)
    return len(current)
