"""
좌석 배치 알고리즘
- 같은 기수끼리 인접 X
- 앞/뒤 편향 최소화
- 이전에 앉았던 자리 피하기
"""
import random
from collections import defaultdict
from pathlib import Path

from reader import Participant, Seat
from history import load_history


def _build_adjacency(seats: list[Seat]) -> dict[str, set[str]]:
    """각 좌석의 인접 좌석 (8방향) 맵"""
    by_pos: dict[tuple[int, int], str] = { (s.row, s.col): s.cell_ref for s in seats }
    adj: dict[str, set[str]] = {}
    for s in seats:
        neighbors = set()
        for dr in (-1, 0, 1):
            for dc in (-1, 0, 1):
                if dr == 0 and dc == 0:
                    continue
                nr, nc = s.row + dr, s.col + dc
                ref = by_pos.get((nr, nc))
                if ref:
                    neighbors.add(ref)
        adj[s.cell_ref] = neighbors
    return adj


def _row_zone(row: int, rows: list[int]) -> str:
    """앞/중간/뒤 구간 (row 숫자가 클수록 교단에 가까움)"""
    if not rows:
        return "middle"
    min_r, max_r = min(rows), max(rows)
    span = max_r - min_r
    if span <= 0:
        return "middle"
    third = span / 3
    if row >= max_r - third:
        return "front"
    if row <= min_r + third:
        return "back"
    return "middle"


def assign(
    seats: list[Seat],
    participants: list[Participant],
    history_path: str | None = None,
    seed: int | None = None,
) -> dict[str, str]:
    """
    참가자를 좌석에 배치.
    반환: { "이름": "A4", ... }
    """
    if len(participants) > len(seats):
        raise ValueError(f"참가자 수({len(participants)})가 좌석 수({len(seats)})보다 많습니다.")

    random.seed(seed)
    seat_refs = [s.cell_ref for s in seats]
    seat_by_ref = {s.cell_ref: s for s in seats}
    adj = _build_adjacency(seats)
    rows = [s.row for s in seats]
    history_list = load_history(Path(history_path) if history_path else None)

    # 사람별 이전에 앉았던 자리
    past_seats: dict[str, list[str]] = defaultdict(list)
    for entry in history_list:
        for name, cell in (entry.get("assignments") or {}).items():
            past_seats[name].append(cell)

    # 사람별 row zone 사용 횟수 (앞/뒤 균형)
    zone_counts: dict[str, dict[str, int]] = defaultdict(lambda: {"front": 0, "middle": 0, "back": 0})
    for entry in history_list:
        for name, cell in (entry.get("assignments") or {}).items():
            seat = seat_by_ref.get(cell)
            if seat:
                zone = _row_zone(seat.row, rows)
                zone_counts[name][zone] = zone_counts[name].get(zone, 0) + 1

    def score(person: Participant, seat_ref: str, assigned: dict[str, str]) -> float:
        """낮을수록 좋은 점수"""
        seat = seat_by_ref[seat_ref]
        penalty = 0.0

        # 1) 같은 기수 인접: 큰 패널티
        for adj_ref in adj.get(seat_ref, set()):
            other_name = assigned.get(adj_ref)
            if other_name:
                other = next((p for p in participants if p.name == other_name), None)
                if other and other.gisu and other.gisu == person.gisu:
                    penalty += 500

        # 2) 이전에 앉았던 자리: 패널티 (가장 최근 것에 더 큰 패널티)
        for i, prev in enumerate(past_seats.get(person.name, [])):
            if prev == seat_ref:
                penalty += 100 * (len(past_seats[person.name]) - i)

        # 3) 앞/뒤 균형: 사용 적은 zone 선호
        zone = _row_zone(seat.row, rows)
        penalty += zone_counts[person.name].get(zone, 0) * 20

        return penalty

    # 셔플 후 그리디 배치 (여러 번 시도해 최선 선택)
    best_assigned: dict[str, str] = {}
    best_penalty = float("inf")

    for attempt in range(30):
        order = participants.copy()
        random.shuffle(order)
        assigned: dict[str, str] = {}
        available = set(seat_refs)
        total_penalty = 0.0

        for person in order:
            best_seat = None
            best_s = float("inf")
            for ref in available:
                s = score(person, ref, assigned)
                if s < best_s:
                    best_s = s
                    best_seat = ref
            if best_seat is None:
                break
            assigned[best_seat] = person.name
            available.remove(best_seat)
            total_penalty += best_s

        if total_penalty < best_penalty:
            best_penalty = total_penalty
            best_assigned = assigned.copy()

    # {이름: 좌석ref} 형식으로 반환 (assignments[seat]=name -> assignments[name]=seat)
    return {name: seat_ref for seat_ref, name in best_assigned.items()}
