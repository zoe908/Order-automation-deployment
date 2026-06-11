"""
범용 PDF 발주서 파서
====================
업체마다 PDF 레이아웃이 조금씩 다르더라도, 대부분의 발주서는
 - 품목 테이블 (품목명/수량/단가 등)
 - 하단 정보 테이블 (납품 장소, 입고 담당자 등 key-value 형태)
 - 상단 텍스트 (발주번호, 발주일자, 납기일자 등)
으로 구성되어 있다는 공통점이 있습니다.

이 모듈은 이 공통 구조를 이용해 PDF에서 데이터를 뽑아내는
범용 함수들을 제공합니다. 업체별로 다른 부분(품목명 텍스트를
어떻게 분해할지 등)은 vendors/ 폴더의 업체별 모듈에서 처리합니다.
"""

import re
import pdfplumber


def extract_pdf_content(pdf_path):
    """PDF에서 모든 테이블과 전체 텍스트를 추출합니다."""
    all_tables = []
    full_text_parts = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text_parts.append(text)
            for table in page.extract_tables():
                all_tables.append(table)

    return all_tables, "\n".join(full_text_parts)


def find_item_table(tables, header_keyword="품목명"):
    """첫 번째 셀이 header_keyword 인 테이블(=품목 테이블)을 찾습니다."""
    for table in tables:
        if table and table[0] and table[0][0] and table[0][0].strip() == header_keyword:
            return table
    return None


def extract_keyvalue_info(tables):
    """
    하단 정보 테이블처럼 ['납품 장소', '주소...', '입고 담당자', '이름...']
    형태로 key-value 가 번갈아 나오는 행들을 모아 dict 로 반환합니다.
    값이 없는(None) 셀은 건너뜁니다.
    """
    info = {}
    for table in tables:
        for row in table:
            cells = [c.strip() if isinstance(c, str) else c for c in row]
            i = 0
            while i < len(cells) - 1:
                key, val = cells[i], cells[i + 1]
                if key and val and key not in ("-",) and val != "-":
                    info[key] = val
                i += 1
    return info


def extract_meta_from_text(full_text, patterns=None):
    """
    전체 텍스트에서 정규식으로 발주번호/발주일자/납기일자 등 메타 정보를 추출합니다.
    patterns: {"필드명": r"정규식(그룹1)"} 형태. 기본값은 한품 계열 발주서 공통 패턴입니다.
    """
    if patterns is None:
        patterns = {
            "발주번호": r"발주번호\s*[:：]?\s*(\S+)",
            "발주일자": r"발주\s*일자\s*[:：]?\s*([0-9]{4}-[0-9]{2}-[0-9]{2})",
            "납기일자": r"납기\s*일자\s*[:：]?\s*([0-9]{4}-[0-9]{2}-[0-9]{2})",
        }

    meta = {}
    for key, pattern in patterns.items():
        m = re.search(pattern, full_text)
        if m:
            meta[key] = m.group(1)
    return meta


def get_item_rows(item_table, name_col=0, skip_header=True):
    """품목 테이블에서 품목명이 비어있지 않은 데이터 행만 반환합니다."""
    rows = item_table[1:] if skip_header else item_table
    return [row for row in rows if row and row[name_col]]
