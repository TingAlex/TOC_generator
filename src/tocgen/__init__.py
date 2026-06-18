"""
tocgen —— PDF 目录识别 · 书签 · 按章节拆分 · OneNote 本地整理。

分层（解耦 + 聚类）：
  · 基础设施：paths（路径）、registry（每书状态 Excel）、bookconfig（Excel 模板/读写）
  · 领域库：  toc（目录模型）、pdf（渲染/书签/切片）
  · 编排：    split（按目录拆分）、onenote.*（COM 建分区/改标题/删附件）
  · 入口：    cli.*（仅 argparse + 打印，调用上面的库）→ console_scripts `toc-*`

Pipeline 1（识别目录→写书签）由 Claude 自身多模态能力完成：toc-claude 渲染目录页，
Claude 看图写出 toc_parsed.txt，toc-claude 再写书签——全程不调用任何外部 AI API。

数据目录（books-todo/ books-done/ books-work/）相对**当前工作目录**解析，
因此所有 `toc-*` 命令都应在项目根目录下运行。
"""

__version__ = "1.0.0"
