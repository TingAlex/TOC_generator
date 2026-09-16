"""OneNote 本地 COM 自动化（仅 Windows + OneNote 桌面版）。

client —— COM 薄封装（读层级/建分区/页 XML/显式同步/改删页）。
common —— OneNote CLI 共享的小工具（默认笔记本名、分区编号解析、范围限定、文件夹助手）。
copy —— 原生整页复制的 XPS/图片只读验证与续跑。
native_copy —— 驱动 OneNote 自身的“移动或复制页”对话框。
"""
