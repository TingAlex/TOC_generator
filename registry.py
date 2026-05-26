import json
from pathlib import Path

WORK_DIR = Path("books-work")


def _state_path(book_name: str) -> Path:
    stem = book_name.removesuffix(".pdf")
    return WORK_DIR / stem / "state.json"


def load() -> dict:
    """扫描所有 books-work/*/state.json，聚合为 {book_name: state} 字典。"""
    registry = {}
    for state_file in sorted(WORK_DIR.glob("*/state.json")):
        try:
            state = json.loads(state_file.read_text(encoding="utf-8"))
            book_name = state_file.parent.name + ".pdf"
            registry[book_name] = state
        except (json.JSONDecodeError, OSError):
            pass
    return registry


def save(registry: dict) -> None:
    """将每本书的状态写回各自的 books-work/{书名}/state.json。"""
    for book_name, state in registry.items():
        path = _state_path(book_name)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            json.dumps(state, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
