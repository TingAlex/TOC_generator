"""
OneNote 桌面版本地 COM 接口薄封装。

通过 `OneNote.Application` COM 对象直接操作**本地缓存**，离线可用，
不走网络 Graph API，符合“关闭同步”的使用场景。

依赖：comtypes（纯 Python，无需编译）。仅 Windows + OneNote 桌面版（Office16）可用。

主要能力：
  - 读取 笔记本 → 分区 → 页 的层级（页按 OneNote 显示顺序）
  - 读/写单页标题
  - 删除页（送 OneNote 回收站，可恢复）
"""

from dataclasses import dataclass, field
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET

# OneNote 2013 schema 命名空间
ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_NS = {"one": ONE_NS}

# Microsoft OneNote 15.0 Type Library（已注册）。用它强制 comtypes 早绑定，
# 否则会退化为晚绑定 IDispatch，OneNote 解析方法名时报 “库没有注册”。
_TYPELIB = ("{0EA692EE-BB50-4E3C-AEF0-356D91732725}", 1, 1)

# HierarchyScope 枚举
HS_SELF = 0
HS_CHILDREN = 1
HS_NOTEBOOKS = 2
HS_SECTIONS = 3
HS_PAGES = 4

# PageInfo 枚举
PI_BASIC = 0

# XMLSchema 枚举
XS_2013 = 2

# OneNote 为空白页显示的占位标题（不区分大小写）
_PLACEHOLDER_TITLES = {"", "无标题页", "无标题", "untitled page", "untitled"}


@dataclass
class Page:
    id: str
    name: str  # OneNote 中的页标题（层级 XML 的 name 属性）
    level: int = 1


@dataclass
class Section:
    id: str
    name: str
    pages: list[Page] = field(default_factory=list)


@dataclass
class Notebook:
    id: str
    name: str
    sections: list[Section] = field(default_factory=list)


class OneNoteClient:
    def __init__(self):
        import comtypes.client
        # 先生成/加载类型库 → 强制早绑定（vtable 调用），避免晚绑定的 “库没有注册”。
        mod = comtypes.client.GetModule(_TYPELIB)
        # OneNote 未运行时会自动拉起
        self._app = comtypes.client.CreateObject(
            "OneNote.Application", interface=mod.IApplication)

    # ── 读取层级 ─────────────────────────────────────────────────────────
    def get_hierarchy(self) -> list[Notebook]:
        """返回所有笔记本（含分区与页，页按显示顺序）。"""
        xml = self._app.GetHierarchy("", HS_PAGES, XS_2013)
        root = ET.fromstring(xml)
        notebooks = []
        for nb_el in root.iter(f"{{{ONE_NS}}}Notebook"):
            nb = Notebook(id=nb_el.get("ID"), name=nb_el.get("name", ""))
            # 分区可能直接挂在 Notebook 下，也可能在 SectionGroup 里，统一递归收集
            for sec_el in nb_el.iter(f"{{{ONE_NS}}}Section"):
                sec = Section(id=sec_el.get("ID"), name=sec_el.get("name", ""))
                for pg_el in sec_el.iter(f"{{{ONE_NS}}}Page"):
                    sec.pages.append(Page(
                        id=pg_el.get("ID"),
                        name=pg_el.get("name", ""),
                        level=int(pg_el.get("pageLevel", "1")),
                    ))
                nb.sections.append(sec)
            notebooks.append(nb)
        return notebooks

    def find_notebook(self, notebooks: list[Notebook], name: str) -> Notebook | None:
        for nb in notebooks:
            if nb.name == name:
                return nb
        return None

    # ── 占位页判定 ───────────────────────────────────────────────────────
    def is_blank_placeholder(self, page: Page) -> bool:
        """
        判定是否为新建分区自动生成的空白无标题占位页：
        标题是占位标题（如“无标题页”）**且** 页内容里没有图片、没有非空文本
        （双保险，避免误删真实首页；打印页必有图片，绝不会被误判）。
        """
        if page.name.strip().lower() not in {t.lower() for t in _PLACEHOLDER_TITLES}:
            return False
        xml = self._app.GetPageContent(page.id, PI_BASIC, XS_2013)
        root = ET.fromstring(xml)
        if root.find(f".//{{{ONE_NS}}}Image") is not None:
            return False
        for t in root.iter(f"{{{ONE_NS}}}T"):
            if (t.text or "").strip():
                return False
        return True

    # ── 写标题 ───────────────────────────────────────────────────────────
    def set_page_title(self, page_id: str, title: str) -> None:
        """提交仅含 Page ID + Title 的最小 XML 来改标题。"""
        safe = escape(title)
        page_xml = (
            f'<?xml version="1.0"?>'
            f'<one:Page xmlns:one="{ONE_NS}" ID="{escape(page_id)}">'
            f'<one:Title><one:OE><one:T>{safe}</one:T></one:OE></one:Title>'
            f'</one:Page>'
        )
        # 参数：xml, dateExpectedLastModified(0.0=不校验修改时间), xsSchema, force=True
        self._app.UpdatePageContent(page_xml, 0.0, XS_2013, True)

    # ── 删页（送回收站） ─────────────────────────────────────────────────
    def delete_page(self, page_id: str) -> None:
        """删除页；deletePermanently=False → 进 OneNote 回收站，可恢复。"""
        # 参数：objectId, dateExpectedLastModified(0.0=不校验), deletePermanently
        self._app.DeleteHierarchy(page_id, 0.0, False)
