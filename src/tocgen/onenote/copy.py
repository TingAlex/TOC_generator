"""OneNote 原生整页复制的只读校验工具。

本模块**不创建页面、不提交 Page XML**。目标页必须由 OneNote 自身的“移动或复制页”
功能生成，才能连同 XPS 打印源一起保留；这里仅计算页面摘要、验证 XPS 与续跑前缀。
"""

from __future__ import annotations

import base64
from dataclasses import dataclass
import hashlib
import math
import time
import xml.etree.ElementTree as ET

from .client import ONE_NS, PI_ALL, OneNoteClient, Page, Section


class CopyValidationError(RuntimeError):
    """源页不适合原生复制，或目标页只读校验失败。"""


@dataclass(frozen=True)
class ImageSignature:
    """不保存大块图片本体的可比较摘要。"""

    sha256: str
    format: str
    x: float
    y: float
    width: float
    height: float


@dataclass(frozen=True)
class PageSignature:
    """打印页的 XPS 后端与预览图片摘要。"""

    images: tuple[ImageSignature, ...]
    xps_files: int
    callbacks: int


def is_online_path(path: str) -> bool:
    """OneNote 层级中的 HTTP(S) path 表示 OneDrive/SharePoint 笔记本。"""
    return path.lower().startswith(("https://", "http://"))


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
            raise CopyValidationError("图片没有内嵌 Data，无法校验打印页。")
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
        raise CopyValidationError("页面没有打印图片。")
    return tuple(signatures)


def page_signature(page_xml: str) -> PageSignature:
    """读取打印页摘要；缺少 XPS/CallbackID 时拒绝，防止把旧栅格页当成完成。"""
    root = ET.fromstring(page_xml)
    xps_files = [child for child in root
                 if child.tag.rsplit("}", 1)[-1] == "XPSFile"]
    callbacks = sum(
        1 for xps in xps_files for child in xps
        if child.tag.rsplit("}", 1)[-1] == "CallbackID"
    )
    if not xps_files or callbacks < len(xps_files):
        raise CopyValidationError(
            "页面缺少 XPSFile/CallbackID，只剩固定栅格预览图；"
            "这类旧版复制页放大后不会重新变清晰。")
    return PageSignature(
        images=image_signatures(page_xml),
        xps_files=len(xps_files),
        callbacks=callbacks,
    )


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


def assert_pages_equal(expected: PageSignature, actual: PageSignature) -> None:
    """目标必须同时保留 XPS 链和全部预览图片。"""
    if (expected.xps_files, expected.callbacks) != (actual.xps_files, actual.callbacks):
        raise CopyValidationError(
            "XPS 打印源结构不一致："
            f"源页 {expected.xps_files}/{expected.callbacks}，"
            f"目标页 {actual.xps_files}/{actual.callbacks}。")
    assert_images_equal(expected.images, actual.images)


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


def verify_resume_prefix(client: OneNoteClient, source: Section, destination: Section,
                         source_signatures: dict[str, PageSignature],
                         *, current_pages: list[Page] | None = None) -> int:
    """目标必须是源页的原生整页复制前缀；返回可从哪个下标续跑。"""
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
        try:
            actual = page_signature(client.get_page_content(right.id, PI_ALL))
            assert_pages_equal(source_signatures[left.id], actual)
        except CopyValidationError as exc:
            raise CopyValidationError(
                f"分区 {source.name!r} 第 {index} 页不是完整的 OneNote 原生副本：{exc}") from exc
    return len(current)
