"""
발주서 변환 웹앱 (Streamlit)
============================
PDF 발주서를 업체별 엑셀 양식으로 자동 변환합니다.
새 업체는 엑셀 양식 캡쳐본 한 장으로 등록할 수 있습니다.

실행:
    streamlit run app.py

배포:
    GitHub에 푸시 후 https://share.streamlit.io 에서 연결
    Streamlit Secrets 에 ANTHROPIC_API_KEY 를 등록하세요.
"""

import io
import os
import re
import sys
import tempfile
from contextlib import redirect_stdout
from pathlib import Path

import pandas as pd
import streamlit as st

# 앱 루트를 sys.path에 추가 (어느 디렉터리에서 실행해도 import 가능하게)
BASE_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(BASE_DIR))
os.chdir(BASE_DIR)

from main import convert
from vendors import get_vendor_registry, save_json_vendor
from core.vendor_setup import analyze_excel_screenshot
from core.github_sync import push_file as github_push

# ── 페이지 설정 ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="발주서 변환기",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.title("📋 발주서 PDF → 엑셀 자동 변환")
st.caption("PDF 발주서를 업체별 지정 엑셀 양식으로 자동 변환합니다.")

tab_convert, tab_register = st.tabs(["📄 발주서 변환", "🏢 새 업체 등록"])


# ── 탭 1: 발주서 변환 ────────────────────────────────────────────────────────
with tab_convert:
    registry = get_vendor_registry()

    if not registry:
        st.info("등록된 업체가 없습니다. '새 업체 등록' 탭에서 업체를 먼저 추가해주세요.")
    else:
        col_vendor, col_pdf = st.columns([1, 2])

        with col_vendor:
            vendor_options = {k: f"{v.VENDOR_NAME}  ({k})" for k, v in registry.items()}
            selected_key = st.selectbox(
                "업체 선택",
                options=list(vendor_options.keys()),
                format_func=lambda k: vendor_options[k],
            )

        with col_pdf:
            uploaded_pdf = st.file_uploader(
                "발주서 PDF 업로드",
                type=["pdf"],
                key="pdf_upload",
                help="변환할 PDF 발주서를 선택하세요.",
            )

        # 새 PDF를 올리면 이전 변환 결과 초기화
        current_pdf_id = uploaded_pdf.name if uploaded_pdf else None
        if st.session_state.get("_last_pdf") != current_pdf_id:
            st.session_state["_last_pdf"] = current_pdf_id
            st.session_state["_excel_bytes"] = None
            st.session_state["_excel_name"] = None
            st.session_state["_warnings"] = []

        if st.button(
            "🔄 변환 시작",
            type="primary",
            disabled=(uploaded_pdf is None),
            use_container_width=False,
        ):
            with st.spinner("PDF 분석 및 엑셀 변환 중..."):
                tmp_pdf = tmp_xlsx = None
                try:
                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
                        f.write(uploaded_pdf.getvalue())
                        tmp_pdf = f.name
                    tmp_xlsx = tmp_pdf.replace(".pdf", ".xlsx")

                    # convert()는 경고 메시지를 stdout에 출력하므로 캡쳐
                    buf = io.StringIO()
                    with redirect_stdout(buf):
                        convert(selected_key, tmp_pdf, tmp_xlsx)
                    stdout_text = buf.getvalue()

                    with open(tmp_xlsx, "rb") as f:
                        excel_bytes = f.read()

                    base = uploaded_pdf.name.rsplit(".", 1)[0]
                    excel_name = f"{base}_{selected_key}.xlsx"

                    st.session_state["_excel_bytes"] = excel_bytes
                    st.session_state["_excel_name"] = excel_name
                    st.session_state["_warnings"] = [
                        line for line in stdout_text.splitlines() if line.strip()
                    ]

                except Exception as e:
                    st.error(f"변환 중 오류가 발생했습니다: {e}")
                finally:
                    for p in [tmp_pdf, tmp_xlsx]:
                        if p:
                            try:
                                os.unlink(p)
                            except OSError:
                                pass

        # 변환 결과 표시
        if st.session_state.get("_excel_bytes"):
            st.success("변환 완료!")

            warnings = st.session_state.get("_warnings", [])
            if warnings:
                with st.expander("⚠️ 경고 메시지 (마스터 상품리스트 미매칭 항목)", expanded=True):
                    for line in warnings:
                        st.text(line)

            st.download_button(
                label="📥 엑셀 다운로드",
                data=st.session_state["_excel_bytes"],
                file_name=st.session_state["_excel_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# ── 탭 2: 새 업체 등록 ──────────────────────────────────────────────────────
with tab_register:
    # 현재 등록된 업체 목록
    st.subheader("현재 등록된 업체")
    registry = get_vendor_registry()
    if registry:
        rows = [[k, v.VENDOR_NAME] for k, v in registry.items()]
        st.dataframe(
            pd.DataFrame(rows, columns=["업체 키", "업체 이름"]),
            hide_index=True,
            use_container_width=False,
        )
    else:
        st.write("등록된 업체가 없습니다.")

    st.divider()
    st.subheader("새 업체 추가")

    # ANTHROPIC_API_KEY 확인
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        try:
            api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
        except Exception:
            pass

    if not api_key:
        st.warning(
            "엑셀 양식 자동 분석에는 **ANTHROPIC_API_KEY** 가 필요합니다.  \n"
            "Streamlit Cloud라면 *Settings → Secrets* 에, 로컬이라면 환경변수로 설정해주세요."
        )

    # GitHub 자동 커밋 설정 확인
    github_token = os.environ.get("GITHUB_TOKEN", "")
    github_repo = os.environ.get("GITHUB_REPO", "")
    try:
        github_token = github_token or st.secrets.get("GITHUB_TOKEN", "")
        github_repo = github_repo or st.secrets.get("GITHUB_REPO", "")
    except Exception:
        pass

    if not github_token or not github_repo:
        st.info(
            "💡 **GitHub 자동 저장 미설정**: 등록한 업체가 앱 재시작 시 초기화될 수 있습니다.  \n"
            "Streamlit Secrets에 `GITHUB_TOKEN` 과 `GITHUB_REPO` 를 추가하면 영구 저장됩니다."
        )

    with st.form("vendor_register_form", clear_on_submit=False):
        col_key, col_name = st.columns(2)
        with col_key:
            new_key = st.text_input(
                "업체 키 (영문 소문자·숫자)",
                placeholder="예: hanssem",
                help="발주서 변환 시 선택하는 식별자입니다. 영문 소문자로 입력하세요.",
            )
        with col_name:
            new_name = st.text_input(
                "업체 이름",
                placeholder="예: 한샘",
            )

        st.markdown("**① 업체 엑셀 양식 캡쳐본 업로드** *(필수)*")
        st.caption(
            "업체가 요청하는 발주/납품 엑셀 양식의 **헤더 행이 보이는 캡쳐 이미지**를 올려주세요. "
            "AI가 컬럼 구조·색상·수식을 자동으로 분석합니다."
        )
        template_img = st.file_uploader(
            "엑셀 양식 이미지 (PNG / JPG)",
            type=["png", "jpg", "jpeg"],
            key="template_img",
        )

        st.markdown("**② 마스터 상품 목록 CSV 업로드** *(선택)*")
        st.caption(
            "품번·색상·사이즈로 내부 상품코드를 조회하는 파일입니다. 없으면 상품코드 컬럼이 빈칸으로 생성됩니다.  \n"
            "필수 컬럼: `품번`, `색상`, `사이즈`, `상품코드`"
        )
        master_csv = st.file_uploader(
            "마스터 상품 목록 CSV (없으면 건너뛰기)",
            type=["csv"],
            key="master_csv",
        )

        submitted = st.form_submit_button("업체 등록", type="primary")

    if submitted:
        errors = []
        if not new_key:
            errors.append("업체 키를 입력하세요.")
        elif not re.match(r"^[a-z0-9_]+$", new_key):
            errors.append("업체 키는 영문 소문자·숫자·언더스코어만 사용할 수 있습니다.")
        elif new_key in get_vendor_registry():
            errors.append(f"'{new_key}' 는 이미 등록된 업체 키입니다.")
        if not new_name:
            errors.append("업체 이름을 입력하세요.")
        if not template_img:
            errors.append("엑셀 양식 캡쳐본을 업로드하세요.")
        if not api_key:
            errors.append("ANTHROPIC_API_KEY가 설정되지 않았습니다.")

        if errors:
            for msg in errors:
                st.error(msg)
        else:
            with st.spinner("AI가 엑셀 양식을 분석 중입니다... (10~20초 소요)"):
                try:
                    # 1) 캡쳐본 분석
                    output_config = analyze_excel_screenshot(
                        template_img.getvalue(),
                        media_type=template_img.type,
                        api_key=api_key,
                    )

                    # 2) 마스터 CSV 저장 (없으면 헤더만 있는 빈 파일 생성)
                    master_dir = BASE_DIR / "master_data"
                    master_dir.mkdir(exist_ok=True)
                    csv_path = master_dir / f"{new_key}_products.csv"
                    if master_csv:
                        csv_path.write_bytes(master_csv.getvalue())
                    else:
                        csv_path.write_text("품번,색상,사이즈,상품코드\n", encoding="utf-8")

                    # 3) 업체 설정 저장 (parse_rules는 기본값 사용)
                    config = {
                        "vendor_name": new_name,
                        "item_table_header": "품목명",
                        "master_products_file": str(csv_path),
                        "parse_rules": {
                            "brand_strip_pattern": "\\[.*?\\]\\s*",
                            "qty_col": 3,
                            "price_col": 5,
                            "color_keywords": [
                                ["라이트베이지", "LIGHTBEIGE"],
                                ["LIGHTBEIGE", "LIGHTBEIGE"],
                                ["스킨", "SKIN"],
                                ["SKIN", "SKIN"],
                                ["아이보리", "IVORY"],
                                ["IVORY", "IVORY"],
                                ["화이트", "WHITE"],
                                ["WHITE", "WHITE"],
                                ["블랙", "BLACK"],
                                ["BLACK", "BLACK"],
                            ],
                            "default_color": "BLACK",
                            "product_name_keywords": [],
                            "default_product_name": "",
                        },
                        "output_config": output_config,
                    }

                    save_json_vendor(new_key, config)

                    # GitHub 자동 커밋 (설정된 경우)
                    if github_token and github_repo:
                        import json as _json
                        json_bytes = _json.dumps(config, ensure_ascii=False, indent=2).encode("utf-8")
                        github_push(
                            token=github_token,
                            repo=github_repo,
                            file_path=f"vendors/configs/{new_key}.json",
                            content_bytes=json_bytes,
                            commit_message=f"feat: 업체 등록 - {new_name} ({new_key})",
                        )
                        if master_csv:
                            github_push(
                                token=github_token,
                                repo=github_repo,
                                file_path=f"master_data/{new_key}_products.csv",
                                content_bytes=master_csv.getvalue(),
                                commit_message=f"feat: 마스터 상품 목록 추가 - {new_key}",
                            )

                    st.success(f"✅ 업체 **{new_name}** (`{new_key}`) 등록 완료!")
                    st.info("'발주서 변환' 탭에서 바로 사용할 수 있습니다.")

                    with st.expander("AI가 분석한 엑셀 양식 구조 확인", expanded=False):
                        st.json(output_config)

                except Exception as e:
                    st.error(f"등록 중 오류가 발생했습니다: {e}")

