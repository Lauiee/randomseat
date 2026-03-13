"""
배치 기록 저장/로드
"""
import json
from datetime import datetime
from pathlib import Path


def _default_path() -> Path:
    return Path(__file__).parent / "seat_history.json"


def load_history(path: Path | None = None) -> list[dict]:
    """
    과거 배치 기록 로드.
    [{ "date": "...", "assignments": {"홍길동": "A4", ...} }, ...]
    """
    p = path or _default_path()
    if not p.exists():
        return []
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def save_history(assignments: dict[str, str], path: Path | None = None) -> None:
    """
    이번 배치 결과를 기록에 추가.
    assignments: { "이름": "A4", ... }
    """
    p = path or _default_path()
    history = load_history(p)
    entry = {
        "date": datetime.now().isoformat(),
        "assignments": assignments,
    }
    history.append(entry)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
