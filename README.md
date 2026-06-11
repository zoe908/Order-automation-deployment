# PDF 발주서 → 업체별 엑셀 양식 자동 변환

받은 PDF 발주서를 업체(거래처)별로 정해진 엑셀 양식으로 자동 변환하는 도구입니다.

## 폴더 구조

```
order_automation/
├── main.py                # 실행 파일
├── core/
│   ├── pdf_parser.py       # PDF에서 표/텍스트 추출 (공통 로직)
│   └── excel_writer.py     # 업체별 양식대로 엑셀 생성 (공통 로직)
├── vendors/
│   ├── __init__.py         # 업체 등록 목록 (VENDOR_REGISTRY)
│   └── hanpoom_comfortlab.py  # "유나이티드보더스(한품)→컴포트랩" 업체 설정
├── master_data/
│   └── comfortlab_products.csv  # 업체별 마스터 상품리스트
├── input/                  # 변환할 PDF를 넣는 폴더
└── output/                 # 변환 결과 엑셀이 저장되는 폴더
```

## 실행 방법

```bash
cd order_automation
pip install pdfplumber openpyxl

python main.py comfortlab "input/[260605] 유나이티드보더스 발주서.pdf"
```

- 첫 번째 인자: 업체키 (`vendors/__init__.py`의 `VENDOR_REGISTRY`에 등록된 키)
- 두 번째 인자: 변환할 PDF 경로
- 세 번째 인자(선택): 출력 엑셀 경로. 생략하면 `output/` 폴더에 자동 저장됩니다.

실행 후, 마스터 상품리스트에서 매칭되지 않은 품목이 있으면 콘솔에 경고로 표시됩니다.
이 경우 `master_data/*.csv`에 해당 상품(품번/색상/사이즈/품명/상품코드)을 추가하면 됩니다.

## 새 업체(거래처) 추가하는 방법

다른 업체의 PDF 발주서 양식이 추가되면, 아래 3단계로 등록할 수 있습니다.

### 1. 마스터 상품리스트 CSV 추가
`master_data/<업체>_products.csv` 파일을 만들고, 해당 업체 기준
"상품코드, 품번, 품명, 색상, 사이즈" 등 필요한 컬럼으로 작성합니다.
(컬럼 구성은 자유롭게 바꿀 수 있으며, `parse_item`에서 사용하는 키만 맞으면 됩니다.)

### 2. 업체 설정 모듈 작성
`vendors/<업체키>.py` 파일을 만듭니다. `vendors/hanpoom_comfortlab.py`를 복사해서
시작하는 것을 추천합니다. 정의해야 할 항목:

- `VENDOR_NAME`: 업체 이름 (로그 출력용)
- `ITEM_TABLE_HEADER`: PDF의 품목 테이블을 찾기 위한 첫 번째 셀 텍스트 (예: `"품목명"`)
- `MASTER_PRODUCTS_FILE`: 위에서 만든 CSV 경로
- `load_master_products()`, `build_code_finder()`: 기본 구현을 그대로 써도 됩니다
- `parse_item(raw_row, code_finder)`: PDF 품목 행 1개를 받아서, 출력 양식에 필요한
  필드들을 dict로 반환. **PDF 품목명 텍스트의 형식이 업체마다 다르므로, 이 함수가
  업체별로 가장 많이 달라지는 부분입니다.**
- `OUTPUT_CONFIG`: 출력 엑셀의 헤더, 색상, 정렬, 숫자 서식, 수식, 합계 컬럼,
  하단 정보 필드 등을 정의하는 딕셔너리. 각 항목의 의미는
  `vendors/hanpoom_comfortlab.py`의 주석을 참고하세요.

### 3. 업체 등록
`vendors/__init__.py`에 새 모듈을 import하고 `VENDOR_REGISTRY`에 키-값으로 추가합니다.

```python
from . import hanpoom_comfortlab
from . import 새업체모듈

VENDOR_REGISTRY = {
    "comfortlab": hanpoom_comfortlab,
    "새업체키": 새업체모듈,
}
```

이제 `python main.py 새업체키 "PDF경로"` 로 실행하면 됩니다.

## 동작 원리 요약

1. `core/pdf_parser.py`가 PDF에서 표(테이블)와 전체 텍스트를 추출합니다.
2. `ITEM_TABLE_HEADER`로 품목 테이블을 찾고, 하단 key-value 정보(납품 장소 등)와
   상단 메타 정보(발주번호 등)도 함께 추출합니다.
3. 업체별 `parse_item()`이 PDF 품목명 텍스트를 분석해 품번/색상/사이즈/품명을
   알아내고, 마스터 상품리스트에서 상품코드를 조회합니다.
4. `core/excel_writer.py`가 `OUTPUT_CONFIG`에 정의된 양식대로 엑셀을 생성합니다
   (헤더 색상, 정렬, 합계 수식, 하단 정보 등 포함).

## 현재 등록된 업체

| 업체키 | 설명 |
|---|---|
| `comfortlab` | 유나이티드보더스(한품)에서 받은 PDF 발주서 → 컴포트랩 내부 발주/입고 양식으로 변환 |
