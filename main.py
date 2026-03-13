#!/usr/bin/env python3
"""
랜덤 좌석 배치 프로그램

사용법:
  python main.py <엑셀파일> [출력이미지경로]

엑셀 구조:
  - 시트1: 좌석 배치 (O 표시된 칸에만 배정)
  - 시트2: 참가자 명단 (이름, 기수 컬럼)
"""
import argparse
from pathlib import Path

from assigner import assign
from history import save_history
from image_gen import generate_image
from reader import read_participants, read_seats


def main() -> None:
    parser = argparse.ArgumentParser(description="랜덤 좌석 배치")
    parser.add_argument("excel", type=str, help="엑셀 파일 경로 (시트1=좌석배치, 시트2=참가자명단)")
    parser.add_argument("-o", "--output", type=str, default=None, help="출력 이미지 경로 (기본: 배치결과_날짜.png)")
    parser.add_argument("--no-history", action="store_true", help="이번 배치를 기록에 저장하지 않음")
    parser.add_argument("--seed", type=int, default=None, help="랜덤 시드 (재현용)")
    args = parser.parse_args()

    excel_path = Path(args.excel)
    if not excel_path.exists():
        print(f"파일 없음: {excel_path}")
        return

    seats, fixed = read_seats(excel_path)
    participants = read_participants(excel_path)
    fixed_names = set(fixed.keys())
    assignable = [p for p in participants if p.name not in fixed_names]
    print(f"배정 가능 좌석 {len(seats)}개, 참가자 {len(assignable)}명, 고정 {len(fixed)}명")

    if len(assignable) > len(seats):
        print("오류: 참가자 수가 좌석 수보다 많습니다.")
        return
    if len(assignable) == 0 and len(fixed) == 0:
        print("오류: 참가자가 없습니다.")
        return

    assignments = assign(seats, assignable, seed=args.seed)
    assignments = {**fixed, **assignments}
    print("배치 완료")

    out = args.output
    if not out:
        from datetime import datetime
        out = excel_path.parent / f"배치결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    generate_image(assignments, excel_path, out)

    if not args.no_history:
        save_history(assignments)
        print("배치 기록 저장됨")


if __name__ == "__main__":
    main()
