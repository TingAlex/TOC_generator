"""toc-onenote-fix —— 修复 OneNote「正在清理上次打开之后的内容」卡死。

复刻 OneFix 的「Fix Relaunch」：强杀 OneNote 进程 + 重启 + 等就绪，**不删任何缓存/数据**
（用户离线编辑、未同步改动只在本地缓存里，删缓存会丢笔记）。打印副本卡住时先跑它解救。

    toc-onenote-fix                 # 杀进程→重启→等就绪
    toc-onenote-fix --ready-timeout 180
"""

import argparse

from ..onenote.fix import fix_relaunch


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--settle", type=float, default=3.0,
                        help="杀进程后、重启前的停顿秒数（默认 3）")
    parser.add_argument("--ready-timeout", type=float, default=120.0,
                        help="等待 OneNote 清理完成并恢复 COM 响应的超时秒数（默认 120）")
    args = parser.parse_args()

    print("修复 OneNote 卡死（杀进程 + 重启 + 等就绪，不删任何文件）：")
    ok = fix_relaunch(settle=args.settle, ready_timeout=args.ready_timeout)
    raise SystemExit(0 if ok else 1)


if __name__ == "__main__":
    main()
