"""
과거 리스트업 이력 관리 모듈.

"과거 10일 이내 리스트업했던 제품이면 제외" 요구사항을 위해, 어떤 상품(productId)을 언제
리스트업했는지 로컬 JSON 파일에 기록해두고, 실행할 때마다 그 기록을 보고 최근 10일 이내 항목을
걸러낸다. 외부 DB 없이 파일 하나로 충분히 처리되는 규모라 이렇게 단순하게 구현했다.
"""

import json
from datetime import date, datetime, timedelta
from pathlib import Path

DEFAULT_HISTORY_PATH = Path(__file__).resolve().parent.parent / "data" / "state" / "listed_products_history.json"
DEDUP_WINDOW_DAYS = 10


def _load(history_path: Path) -> dict:
    """이력 파일을 읽는다. 파일이 없으면 빈 이력으로 시작한다."""
    if not history_path.exists():
        return {}
    with open(history_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _save(history_path: Path, history: dict) -> None:
    history_path.parent.mkdir(parents=True, exist_ok=True)
    with open(history_path, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


def filter_out_recent(product_ids: list[str], history_path: Path = DEFAULT_HISTORY_PATH,
                       window_days: int = DEDUP_WINDOW_DAYS, today: date | None = None) -> list[str]:
    """최근 window_days일 이내에 이미 리스트업된 productId를 제외하고 반환한다."""
    today = today or date.today()
    history = _load(history_path)
    cutoff = today - timedelta(days=window_days)

    kept = []
    for pid in product_ids:
        last_listed = history.get(pid)
        if last_listed is None:
            kept.append(pid)
            continue
        last_date = datetime.strptime(last_listed, "%Y-%m-%d").date()
        if last_date < cutoff:
            kept.append(pid)
    return kept


def record_listed(product_ids: list[str], history_path: Path = DEFAULT_HISTORY_PATH,
                   today: date | None = None) -> None:
    """이번에 새로 리스트업한 productId들을 오늘 날짜로 이력에 기록한다."""
    today = today or date.today()
    history = _load(history_path)
    for pid in product_ids:
        history[pid] = today.strftime("%Y-%m-%d")
    _save(history_path, history)


if __name__ == "__main__":
    import tempfile
    test_path = Path(tempfile.gettempdir()) / "listing_history_test.json"
    test_path.unlink(missing_ok=True)

    print("1차 실행 (이력 없음):", filter_out_recent(["A", "B", "C"], test_path))
    record_listed(["A", "B", "C"], test_path)
    print("2차 실행 (방금 등록됨, 전부 제외되어야 함):", filter_out_recent(["A", "B", "C", "D"], test_path))

    old_date = date.today() - timedelta(days=11)
    record_listed(["A"], test_path, today=old_date)
    print("A만 11일 전으로 재기록 후 (A는 다시 나와야 함):", filter_out_recent(["A", "B", "D"], test_path))
    test_path.unlink(missing_ok=True)
