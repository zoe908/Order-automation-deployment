"""
업체: 유나이티드보더스(한품) -> 컴포트랩 발주서 변환 설정

한품에서 보내주는 "[업체명] 품목명 - 색상 (사이즈) 품번" 형태의
PDF 발주서를 컴포트랩 내부 발주/입고 양식(상품코드/품번/품명/색상/사이즈/ORDER/단가...)
으로 변환합니다.
"""

import csv
import os
import re

VENDOR_NAME = "유나이티드보더스(한품)"

# 품목 테이블을 찾기 위한 헤더 키워드 (PDF 표의 첫 번째 셀)
ITEM_TABLE_HEADER = "품목명"

# 마스터 상품리스트 CSV 경로 (이 파일 기준 상대경로)
MASTER_PRODUCTS_FILE = os.path.join(
    os.path.dirname(__file__), "..", "master_data", "comfortlab_products.csv"
)

# 품목명 텍스트에서 색상을 판별하기 위한 키워드.
# 위에서부터 순서대로 검사하며, 일치하는 것이 없으면 기본값(BLACK)을 사용합니다.
COLOR_KEYWORDS = [
    ("라이트베이지", "LIGHTBEIGE"),
    ("LIGHTBEIGE", "LIGHTBEIGE"),
    ("스킨", "SKIN"),
    ("SKIN", "SKIN"),
    ("아이보리", "IVORY"),
    ("IVORY", "IVORY"),
    ("화이트", "WHITE"),
    ("WHITE", "WHITE"),
    ("블랙", "BLACK"),
    ("BLACK", "BLACK"),
]
DEFAULT_COLOR = "BLACK"

# 품목명 텍스트에서 품명을 판별하기 위한 키워드.
# (텍스트에 포함된 키워드 -> 컴포트랩 내부 품명)
PRODUCT_NAME_KEYWORDS = [
    ("팬티", "마카롱웰핏팬티"),
    ("튜브탑", "듀얼쿨튜브탑"),
]
DEFAULT_PRODUCT_NAME = "마카롱웰핏브라렛"


def load_master_products():
    """마스터 상품리스트 CSV를 읽어 dict 리스트로 반환합니다."""
    with open(MASTER_PRODUCTS_FILE, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def build_code_finder(master_products):
    """(품번, 색상, 사이즈) 조합 -> 상품코드 매핑 딕셔너리를 생성합니다."""
    return {
        (p["품번"], p["색상"], p["사이즈"]): p["상품코드"]
        for p in master_products
    }


def parse_item(raw_row, code_finder):
    """
    PDF 품목 테이블의 한 행을 받아 컴포트랩 내부 양식에 필요한 dict 를 반환합니다.

    raw_row 예시:
      ['[컴포트랩] 마카롱 웰핏 브라렛 - 블랙 (95) BR221', '과세', 'ea', '5', '5', '12537', '56986', '5699']
       품목명                                              구분    단위   주문수량  총낱개  개별단가   공급가액   부가세액
    """
    full_text = raw_row[0].replace("\n", " ").strip()
    qty = int(raw_row[3])
    price = int(raw_row[5])

    # 1. 대괄호로 표시된 브랜드명 제거 (예: "[컴포트랩] ")
    clean_text = re.sub(r"^\[.*?\]\s*", "", full_text)

    # 2. 품번 추출: 텍스트의 마지막 단어 (예: BR221, PT221, BR1034)
    item_no = clean_text.split()[-1]

    # 3. 사이즈 추출: 괄호 안의 값. 괄호가 없으면 마지막에서 두 번째 단어 사용
    #    (예: "...팬티 - 블랙 XL PT221" -> "XL")
    size_match = re.search(r"\((.*?)\)", clean_text)
    if size_match:
        size = size_match.group(1)
    else:
        tokens = clean_text.split()
        size = tokens[-2] if len(tokens) >= 2 else ""

    # 4. 색상 판별
    upper_text = clean_text.upper()
    color = DEFAULT_COLOR
    for keyword, mapped_color in COLOR_KEYWORDS:
        if keyword in clean_text or keyword in upper_text:
            color = mapped_color
            break

    # 5. 품명 판별
    product_name = DEFAULT_PRODUCT_NAME
    for keyword, mapped_name in PRODUCT_NAME_KEYWORDS:
        if keyword in clean_text:
            product_name = mapped_name
            break

    # 6. 마스터 데이터에서 상품코드 조회
    product_code = code_finder.get((item_no, color, size), "코드없음")

    return {
        "상품코드": product_code,
        "품번": item_no,
        "품명": product_name,
        "색상": color,
        "사이즈": size,
        "ORDER": qty,
        "단가": price,
    }


# ==========================================
# 출력 엑셀 양식 정의
# ==========================================
OUTPUT_CONFIG = {
    "sheet_name": "발주서양식",
    "master_sheet_name": "제품리스트",
    "headers": ["상품코드", "품번", "품명", "색상", "사이즈", "ORDER", "단가", "단가*수량", "재고", "비고"],
    "header_fills": {
        "상품코드": "FFF2CC",
        "품번": "FFF2CC",
        "품명": "FFF2CC",
        "ORDER": "FFFF00",
        "단가": "FFFF00",
    },
    # 컬럼 인덱스(1부터 시작) -> 정렬
    "column_alignment": {
        1: "center", 2: "center", 4: "center", 5: "center",  # 상품코드/품번/색상/사이즈
        6: "right", 7: "right", 8: "right",                    # ORDER/단가/단가*수량
    },
    # 컬럼 인덱스 -> 숫자 서식
    "number_format_columns": {
        6: "#,##0",
        7: "₩ #,##0",
        8: "₩ #,##0",
    },
    # 컬럼 인덱스 -> 수식 (행 번호는 {row} 로 치환됨). 8번째 컬럼 = 단가*수량 = F*G
    "formula_columns": {
        8: "=F{row}*G{row}",
    },
    # 합계(SUM) 행을 만들 컬럼 (ORDER 합계, 단가*수량 합계)
    "sum_columns": [6, 8],
    "footer_width": 5,
    # PDF 하단 정보 테이블에서 그대로 가져와 표시할 필드
    "footer_fields": [
        {"label": "납품 장소", "pdf_key": "납품 장소"},
        {"label": "입고 담당자", "pdf_key": "입고 담당자"},
    ],
}
