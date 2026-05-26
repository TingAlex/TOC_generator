"""快速测试：只处理 books-todo 里的第一本书。

用法：
    uv run python test_one.py           # dry-run，只跑识别，不写 PDF
    uv run python test_one.py --write   # 写入 PDF
"""

import argparse
from dotenv import load_dotenv
load_dotenv()

from pathlib import Path
import registry as reg
from main import process_one

parser = argparse.ArgumentParser()
parser.add_argument("--write", action="store_true", help="写入 PDF（默认 dry-run）")
args = parser.parse_args()

todo_dir = Path("books-todo")
done_dir = Path("books-done")
done_dir.mkdir(exist_ok=True)

pdfs = sorted(todo_dir.glob("*.pdf"))
if not pdfs:
    print("books-todo/ 中没有 PDF 文件。")
else:
    first = pdfs[0]
    registry = reg.load()
    print(f"测试文件：{first.name}")
    print(f"模式：{'写入 PDF' if args.write else 'dry-run（不写 PDF）'}\n")
    process_one(first, done_dir, registry, write=args.write)
