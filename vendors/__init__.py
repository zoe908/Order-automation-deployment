"""
업체별(거래처별) 발주서 변환 설정 모듈 모음.

업체 추가 방법:
  - 웹 UI: '새 업체 등록' 탭에서 엑셀 양식 캡쳐본 + 마스터 상품 목록 CSV 업로드
  - 수동: vendors/configs/<업체키>.json 파일 작성 (hanpoom_comfortlab.py 참고)
  - Python 모듈: 기존 방식 그대로 사용 가능
"""

import re
import csv
import json
from pathlib import Path

from . import hanpoom_comfortlab

_CONFIGS_DIR = Path(__file__).parent / "configs"


class VendorFromConfig:
    """vendors/configs/*.json 파일로부터 동적으로 생성되는 업체 모듈 객체"""

    def __init__(self, config: dict, config_path: str = None):
        self.VENDOR_NAME = config["vendor_name"]
        self.ITEM_TABLE_HEADER = config.get("item_table_header", "품목명")
        self._parse_rules = config.get("parse_rules", {})
        self._master_file = config.get("master_products_file", "")
        self._config_path = config_path

        # JSON 키는 문자열이므로 숫자 인덱스 필드를 int로 변환
        oc = dict(config.get("output_config", {}))
        for field in ("column_alignment", "number_format_columns", "formula_columns"):
            if field in oc and isinstance(oc[field], dict):
                oc[field] = {int(k): v for k, v in oc[field].items()}
        if "sum_columns" in oc:
            oc["sum_columns"] = [int(x) for x in oc["sum_columns"]]
        self.OUTPUT_CONFIG = oc

    # ── 마스터 상품리스트 ──────────────────────────────────────────────────

    def load_master_products(self):
        with open(self._master_file, encoding="utf-8-sig") as f:
            return list(csv.DictReader(f))

    def build_code_finder(self, master_products):
        return {
            (p.get("품번", ""), p.get("색상", ""), p.get("사이즈", "")): p.get("상품코드", "")
            for p in master_products
        }

    # ── PDF 품목 행 파싱 ──────────────────────────────────────────────────

    def parse_item(self, raw_row, code_finder):
        rules = self._parse_rules
        full_text = str(raw_row[0] or "").replace("\n", " ").strip()

        qty_col = rules.get("qty_col", 3)
        price_col = rules.get("price_col", 5)

        def to_int(val):
            try:
                return int(str(val).replace(",", "").strip())
            except (ValueError, TypeError):
                return 0

        qty = to_int(raw_row[qty_col]) if len(raw_row) > qty_col else 0
        price = to_int(raw_row[price_col]) if len(raw_row) > price_col else 0

        # 브랜드 접두어 제거 (예: "[컴포트랩] ")
        brand_pattern = rules.get("brand_strip_pattern")
        clean = re.sub(brand_pattern, "", full_text) if brand_pattern else full_text
        tokens = clean.split()

        # 품번: 마지막 토큰
        item_no = tokens[-1] if tokens else ""

        # 사이즈: 괄호 안 또는 뒤에서 두 번째 토큰
        size_match = re.search(r"\(([^)]+)\)", clean)
        size = size_match.group(1) if size_match else (tokens[-2] if len(tokens) >= 2 else "")

        # 색상: 키워드 순서 검사
        upper = clean.upper()
        color = rules.get("default_color", "")
        for pair in rules.get("color_keywords", []):
            kw, mapped = pair[0], pair[1]
            if kw in clean or kw.upper() in upper:
                color = mapped
                break

        # 품명: 키워드 순서 검사
        product_name = rules.get("default_product_name", "")
        for pair in rules.get("product_name_keywords", []):
            kw, mapped = pair[0], pair[1]
            if kw in clean:
                product_name = mapped
                break

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


# ── 레지스트리 로딩 ─────────────────────────────────────────────────────────

def _load_json_vendors() -> dict:
    """vendors/configs/*.json 파일을 읽어 VendorFromConfig 객체 딕셔너리로 반환합니다."""
    if not _CONFIGS_DIR.exists():
        return {}
    vendors = {}
    for cfg_file in sorted(_CONFIGS_DIR.glob("*.json")):
        try:
            config = json.loads(cfg_file.read_text(encoding="utf-8"))
            vendors[cfg_file.stem] = VendorFromConfig(config, str(cfg_file))
        except Exception as e:
            print(f"[경고] 업체 설정 로드 실패 ({cfg_file.name}): {e}")
    return vendors


# CLI(main.py)용 모듈 수준 레지스트리
VENDOR_REGISTRY: dict = {
    "comfortlab": hanpoom_comfortlab,
    **_load_json_vendors(),
}


def get_vendor_registry() -> dict:
    """최신 업체 목록을 반환합니다. JSON 파일 변경사항(신규 등록 등)도 반영됩니다."""
    return {
        "comfortlab": hanpoom_comfortlab,
        **_load_json_vendors(),
    }


def save_json_vendor(vendor_key: str, config: dict) -> None:
    """새 업체 설정을 JSON 파일로 저장하고 모듈 수준 VENDOR_REGISTRY도 갱신합니다."""
    _CONFIGS_DIR.mkdir(exist_ok=True)
    cfg_path = _CONFIGS_DIR / f"{vendor_key}.json"
    cfg_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    VENDOR_REGISTRY[vendor_key] = VendorFromConfig(config, str(cfg_path))
