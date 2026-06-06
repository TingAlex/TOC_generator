"""命令行入口（薄壳）：仅 argparse + 打印，业务逻辑在 tocgen 的库模块里。

每个模块暴露 main()，对应 pyproject [project.scripts] 的一个 toc-* 命令。
"""
