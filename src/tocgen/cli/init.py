"""toc-init —— 扫描 books-todo/(与 books-done/) 登记新书到 Excel，并生成 split_config.xlsx。

新书入库后重跑即可（幂等，不动既有行）。
"""

from .. import bookconfig


def main() -> None:
    bookconfig.run_init()


if __name__ == "__main__":
    main()
