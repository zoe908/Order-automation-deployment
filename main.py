"""
PDF 발주서 -> 업체별 엑셀 양식 변환 자동화

사용법:
    python main.py <업체키> <PDF경로> [출력엑셀경로]

예시:
    python main.py comfortlab "input/[260605] 유나이티드보더스 발주서.pdf"
    python main.py comfortlab "input/발주서.pdf" "output/결과.xlsx"

업체키 목록은 vendors/__init__.py 의 VENDOR_REGISTRY 를 참고하세요.
새 업체를 추가하는 방법은 README.md 를 참고하세요.
"""

import os
import sys

from core.pdf_parser import (
    extract_pdf_content,
    find_item_table,
    extract_keyvalue_info,
    extract_meta_from_text,
    get_item_rows,
)
from core.excel_writer import write_order_excel
from vendors import get_vendor_registry


def convert(vendor_key, pdf_path, output_path=None):
    vendor_registry = get_vendor_registry()
    if vendor_key not in vendor_registry:
        available = ", ".join(vendor_registry.keys())
        raise ValueError(f"알 수 없는 업체키 '{vendor_key}'. 사용 가능: {available}")

    vendor = vendor_registry[vendor_key]

    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")

    if output_path is None:
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        output_path = os.path.join("output", f"{base}_{vendor_key}_변환.xlsx")
        os.makedirs("output", exist_ok=True)

    # 1) PDF에서 표/텍스트 추출
    tables, full_text = extract_pdf_content(pdf_path)

    # 2) 품목 테이블 찾기
    item_table = find_item_table(tables, vendor.ITEM_TABLE_HEADER)
    if item_table is None:
        raise ValueError(
            f"PDF에서 품목 테이블('{vendor.ITEM_TABLE_HEADER}' 헤더)을 찾지 못했습니다. "
            f"PDF 레이아웃을 확인하거나 vendors/{vendor_key}.py 의 ITEM_TABLE_HEADER 설정을 확인하세요."
        )
    raw_rows = get_item_rows(item_table)

    # 3) 마스터 상품리스트 로드 + 품목 파싱/매핑
    master_products = vendor.load_master_products()
    code_finder = vendor.build_code_finder(master_products)
    data_rows = [vendor.parse_item(row, code_finder) for row in raw_rows]

    # 4) 코드 매칭 실패 항목 경고
    unmatched = [r for r in data_rows if r["상품코드"] == "코드없음"]
    if unmatched:
        print("[경고] 마스터 상품리스트에서 매칭되지 않은 품목이 있습니다:")
        for r in unmatched:
            print(f"   - 품번={r['품번']}, 품명={r['품명']}, 색상={r['색상']}, 사이즈={r['사이즈']}")
        print("   -> master_data 의 CSV에 해당 상품을 추가해주세요.\n")

    # 5) 하단 정보(납품 장소 등) + 메타 정보(발주번호 등) 추출
    footer_info = extract_keyvalue_info(tables)
    meta = extract_meta_from_text(full_text)

    # 6) 엑셀 생성
    write_order_excel(
        output_path=output_path,
        data_rows=data_rows,
        output_config=vendor.OUTPUT_CONFIG,
        footer_info=footer_info,
        meta=meta,
        master_products=master_products,
    )

    print(f"완료: '{output_path}' 생성됨 ({len(data_rows)}개 품목, 업체: {vendor.VENDOR_NAME})")
    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    vendor_key = sys.argv[1]
    pdf_path = sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) > 3 else None

    convert(vendor_key, pdf_path, output_path)
