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

# CreateFileType 枚举（OpenHierarchy 的 cftIfNotExist：不存在时创建何种对象）
CFT_NONE = 0
CFT_NOTEBOOK = 1
CFT_FOLDER = 2   # 分区组（section group，磁盘上是文件夹）
CFT_SECTION = 3  # 分区（.one 文件）

# SetFilingLocation 枚举：把「print to OneNote 打印机」的打印输出定向到指定分区。
# FilingLocation.flPrintOuts=5（打印机输出），FilingLocationType.fltNamedSectionNewPage=0
# （在指定分区里新建一页归档）。设好后该分区即成为打印输出落点，无需弹位置选择框。
FL_PRINTOUTS = 5
FLT_NAMED_SECTION_NEW_PAGE = 0

# OneNote 为空白页显示的占位标题（不区分大小写）
_PLACEHOLDER_TITLES = {"", "无标题页", "无标题", "untitled page", "untitled"}


@dataclass
class Page:
    id: str
    name: str  # OneNote 中的页标题（层级 XML 的 name 属性）
    level: int = 1


@dataclass
class InsertedFile:
    """页内作为附件嵌入的源文件（OneNote Batch 导入时连同打印图片一起塞进来的）。"""
    object_id: str    # 可删除对象的 objectID（最近一个带 objectID 的祖先 OE）
    name: str         # preferredName，如 "117-4.4.1.pdf"
    path_source: str  # pathSource，导入时的源文件磁盘路径（可能为空）


@dataclass
class Section:
    id: str
    name: str
    pages: list[Page] = field(default_factory=list)


@dataclass
class SectionGroup:
    id: str
    name: str
    sections: list[Section] = field(default_factory=list)          # 组内直属分区
    section_groups: list["SectionGroup"] = field(default_factory=list)  # 嵌套子分区组


@dataclass
class Notebook:
    id: str
    name: str
    path: str = ""  # 笔记本位置：在线（OneDrive）笔记本以 https:// 开头，本地笔记本是磁盘路径
    # sections：扁平递归列表，含直属分区 + 各分区组内所有分区（向后兼容，按名查找用）
    sections: list[Section] = field(default_factory=list)
    section_groups: list[SectionGroup] = field(default_factory=list)  # 直属分区组（嵌套树）


def _flatten_groups(groups: list[SectionGroup]) -> list[Section]:
    """递归收集所有分区组内的分区，拼成扁平列表。"""
    out: list[Section] = []
    for g in groups:
        out.extend(g.sections)
        out.extend(_flatten_groups(g.section_groups))
    return out


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
            nb = Notebook(id=nb_el.get("ID"), name=nb_el.get("name", ""),
                          path=nb_el.get("path", ""))
            # 嵌套树（直属分区 + 直属分区组，分区组内再递归）
            direct, nb.section_groups = self._parse_container(nb_el)
            # 扁平递归列表：直属 + 各组内所有分区（向后兼容）
            nb.sections = direct + _flatten_groups(nb.section_groups)
            notebooks.append(nb)
        return notebooks

    def _parse_section(self, sec_el) -> Section:
        sec = Section(id=sec_el.get("ID"), name=sec_el.get("name", ""))
        for pg_el in sec_el.iter(f"{{{ONE_NS}}}Page"):
            sec.pages.append(Page(
                id=pg_el.get("ID"),
                name=pg_el.get("name", ""),
                level=int(pg_el.get("pageLevel", "1")),
            ))
        return sec

    def _parse_container(self, el) -> tuple[list[Section], list[SectionGroup]]:
        """解析容器（Notebook 或 SectionGroup）的**直属**分区与分区组（分区组内递归）。"""
        sections = [self._parse_section(s)
                    for s in el.findall(f"{{{ONE_NS}}}Section")]
        groups = []
        for sg_el in el.findall(f"{{{ONE_NS}}}SectionGroup"):
            # 跳过回收站（isRecycleBin="true"）：里面是已删分区，不应纳入处理范围
            if sg_el.get("isRecycleBin", "").lower() == "true":
                continue
            sg = SectionGroup(id=sg_el.get("ID"), name=sg_el.get("name", ""))
            sg.sections, sg.section_groups = self._parse_container(sg_el)
            groups.append(sg)
        return sections, groups

    def find_section_group(self, container, name: str) -> SectionGroup | None:
        """在笔记本/分区组下按名递归查找分区组。container 需有 .section_groups。"""
        for g in container.section_groups:
            if g.name == name:
                return g
            sub = self.find_section_group(g, name)
            if sub is not None:
                return sub
        return None

    def find_notebook(self, notebooks: list[Notebook], name: str) -> Notebook | None:
        for nb in notebooks:
            if nb.name == name:
                return nb
        return None

    def list_section_pages(self, section_id: str) -> list[Page]:
        """只取某个分区的页（按显示顺序）。比全量 get_hierarchy 轻，供打印落地轮询用。"""
        xml = self._app.GetHierarchy(section_id, HS_PAGES, XS_2013)
        sec_el = ET.fromstring(xml)
        return self._parse_section(sec_el).pages

    # ── 打印导入：把打印输出定向到指定分区 ──────────────────────────────
    def set_printout_section(self, section_id: str) -> None:
        """把「print to OneNote 打印机」的打印输出落点设为该分区（之后每次打印新建一页）。

        持久改 OneNote 的打印输出归档设置；无读回接口，跑完不还原（仅影响用户下次手动打印的默认落点）。
        """
        self._app.SetFilingLocation(FL_PRINTOUTS, FLT_NAMED_SECTION_NEW_PAGE, section_id)

    # ── 创建层级（分区组 / 分区 / 在线笔记本） ───────────────────────────
    def create_section_group(self, notebook_id: str, name: str) -> str:
        """在笔记本下新建分区组，返回其 ID。路径相对于 notebook_id。"""
        # OpenHierarchy(path, relativeToObjectID, [out]objectID, cftIfNotExist)
        # comtypes 早绑定把 [out] objectID 转成返回值。
        return self._app.OpenHierarchy(name, notebook_id, CFT_FOLDER)

    def create_section(self, parent_id: str, name: str) -> str:
        """
        在父对象（笔记本或分区组）下新建分区，返回其 ID。

        分区在磁盘上是 `.one` 文件，OpenHierarchy 的路径**必须带 `.one` 扩展名**
        （分区组/文件夹则不带）；OneNote 显示时会去掉该扩展名。
        """
        path = name if name.lower().endswith(".one") else f"{name}.one"
        return self._app.OpenHierarchy(path, parent_id, CFT_SECTION)

    def create_local_notebook(self, local_path: str) -> str:
        """新建本地（不同步）笔记本，返回其 ID。local_path 为磁盘绝对路径（文件夹名即笔记本名）。"""
        return self._app.OpenHierarchy(local_path, "", CFT_NOTEBOOK)

    def create_online_notebook(self, name: str, sibling_of_path: str) -> str:
        """
        新建**在线**笔记本，做成现有在线笔记本的同级。

        sibling_of_path 为参考在线笔记本的 path（OneDrive URL，形如
        `https://d.docs.live.net/<id>/文档/某本子`）。取其父目录拼上 name 作为新笔记本路径，
        从而落在同一 OneDrive 位置（联网笔记本），而非本地。
        """
        # 用 URL 的 '/' 分隔，绝不能用本地 os.path（会引入反斜杠）。
        parent = sibling_of_path.rstrip("/").rsplit("/", 1)[0]
        new_path = f"{parent}/{name}"
        return self._app.OpenHierarchy(new_path, "", CFT_NOTEBOOK)

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

    # ── 嵌入附件：列出 / 删除 ────────────────────────────────────────────
    def list_inserted_files(self, page_id: str,
                            exts: set[str] | None = None) -> list[InsertedFile]:
        """
        列出页内可删除的嵌入文件附件（默认只挑 .pdf 源文件）。

        OneNote 页 XML 里附件是 `<one:InsertedFile preferredName="x.pdf" pathSource="...">`，
        它本身一般不带 objectID —— objectID 在外层 `<one:OE>` 上。这里自附件向上找最近一个
        带 objectID 的祖先 OE 作为删除目标；若该祖先子树里含 `<one:Image>`（打印图片），则
        跳过（绝不连带删图片），保证只删纯附件。
        """
        if exts is None:
            exts = {".pdf"}
        exts = {e.lower() if e.startswith(".") else "." + e.lower() for e in exts}

        xml = self._app.GetPageContent(page_id, PI_BASIC, XS_2013)
        root = ET.fromstring(xml)
        # ElementTree 无父指针，先建子→父映射
        parent = {child: el for el in root.iter() for child in el}

        results: list[InsertedFile] = []
        for f in root.iter(f"{{{ONE_NS}}}InsertedFile"):
            name = f.get("preferredName", "")
            src = f.get("pathSource", "")
            ext_src = (name or src).lower()
            if not any(ext_src.endswith(e) for e in exts):
                continue
            # 向上找最近带 objectID 的祖先（含自身）
            node = f
            obj_id = None
            while node is not None:
                if node.get("objectID"):
                    obj_id = node.get("objectID")
                    break
                node = parent.get(node)
            if obj_id is None:
                continue  # 没有可删除的对象 ID，跳过
            # 安全护栏：删除目标子树内若含图片则放弃（不误删打印页）
            if node.find(f".//{{{ONE_NS}}}Image") is not None:
                continue
            results.append(InsertedFile(object_id=obj_id, name=name, path_source=src))
        return results

    def delete_page_content(self, page_id: str, object_id: str) -> None:
        """删除页内某个对象（如嵌入文件附件）。"""
        # 参数：pageId, objectId, dateExpectedLastModified(0.0=不校验), force=True
        self._app.DeletePageContent(page_id, object_id, 0.0, True)

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
