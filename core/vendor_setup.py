"""
Claude Vision을 이용해 엑셀 양식 이미지를 분석하고
OUTPUT_CONFIG 딕셔너리를 자동 생성합니다.
"""

import anthropic
import base64
import json
import re


def analyze_excel_screenshot(image_bytes: bytes, media_type: str = "image/png", api_key: str = None) -> dict:
    """
    엑셀 양식 이미지를 Claude Vision으로 분석하여 OUTPUT_CONFIG를 반환합니다.

    Returns:
        dict: excel_writer.write_order_excel()에서 사용하는 output_config 딕셔너리
    """
    if not media_type or not media_type.startswith("image/"):
        media_type = "image/png"

    client = anthropic.Anthropic(api_key=api_key)
    image_b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    prompt = """이 이미지는 발주서 엑셀 양식입니다.

양식의 구조를 분석해서 아래 JSON 형식으로만 응답하세요.
마크다운 코드블록(```) 없이, 순수 JSON만 출력하세요.

{
  "sheet_name": "시트 이름 (보이면 사용, 없으면 '발주서')",
  "master_sheet_name": "제품리스트",
  "headers": ["컬럼1", "컬럼2", ...],
  "header_fills": {
    "배경색 있는 헤더명": "RRGGBB (예: FFF2CC, FFFF00)"
  },
  "column_alignment": {
    "컬럼번호문자열": "center 또는 right 또는 left"
  },
  "number_format_columns": {
    "컬럼번호문자열": "#,##0 또는 ₩ #,##0"
  },
  "formula_columns": {
    "컬럼번호문자열": "=C{row}*D{row} 형식"
  },
  "sum_columns": [합계가 필요한 컬럼번호 정수 배열],
  "footer_fields": [],
  "footer_width": 5
}

규칙:
- headers: 이미지에서 보이는 컬럼명을 순서대로 정확히 기재
- header_fills: 배경색이 있는 헤더만 포함 (색상 없으면 빈 객체 {})
- column_alignment: 컬럼번호는 1부터 시작하는 문자열 키
- number_format_columns: 숫자·금액 컬럼 번호 (문자열 키)
- formula_columns: 다른 셀 값 연산이 필요한 컬럼 ({row}는 실제 행번호로 치환됨)
- sum_columns: 합계행이 필요한 컬럼의 번호 (정수 배열)
- 컬럼명은 이미지에 보이는 그대로 정확히 사용하세요"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": image_b64,
                        },
                    },
                    {"type": "text", "text": prompt},
                ],
            }
        ],
    )

    raw = response.content[0].text.strip()

    # 마크다운 코드펜스 제거
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    raw = raw.strip()

    config = json.loads(raw)

    # sum_columns를 정수 배열로 정규화
    if "sum_columns" in config:
        config["sum_columns"] = [int(x) for x in config["sum_columns"]]

    return config
