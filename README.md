# randomseat

엑셀 파일을 넣으면 좌석을 랜덤 배치하고 결과를 이미지로 저장하는 프로그램입니다.

## 설치

```bash
pip install -r requirements.txt
```

## 엑셀 형식

- **시트 1 (좌석배치)**: 좌석이 있는 칸에 `O` 표시
- **시트 2 (참가자명단)**: 첫 행에 `이름`, `기수` 컬럼

## 실행

```bash
python main.py <엑셀파일> [-o 출력이미지.png]
```

**옵션**

- `-o`, `--output` : 출력 이미지 경로 (기본: `배치결과_날짜시간.png`)
- `--no-history` : 이번 배치를 기록에 저장하지 않음
- `--seed N` : 랜덤 시드 (같은 결과 재현용)

## 규칙

1. **같은 기수끼리 인접하지 않도록** 분산 배치
2. **앞/뒤 편향 없이** 골고루 배치
3. **과거 배치 기록**을 저장해, 이전에 앉았던 자리를 피해 배치

## 테스트

```bash
python create_sample.py   # 샘플 엑셀 생성
python main.py sample.xlsx
```

배치 기록은 `seat_history.json`에 저장됩니다.
