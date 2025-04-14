import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

# pyhwpx가 설치되어 있으면 import
# 윈도우 이외 OS 동작오류 방지
try:
    from hwp.service import hwp_context, open_template, process_documents
except ImportError:
    pass


def main():
    st.set_page_config(page_title="AutoHWP", layout="wide")
    st.markdown(
        """
        <style>
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');

        html, body, [class*="css"]  {
            font-family: 'Pretendard', sans-serif !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("양식문서 자동완성")

    # 워크플로우 이름
    workflow_name = st.text_input(
        "워크플로우 이름",
        value=st.session_state.get("workflow_name", ""),
        placeholder="예: 20XX_양식문서",
        max_chars=50,
        help="워크플로우 이름은 생성된 파일명의 접두사로 사용됩니다. (예: 20XX_양식문서_홍길동.hwp)",
        # label_visibility="collapsed",
    )
    if workflow_name:
        st.session_state["workflow_name"] = workflow_name

    output_folder = st.text_input(
        "출력 폴더",
        value=Path(tempfile.gettempdir()) / (workflow_name or "autohwp"),
        placeholder="예: /tmp/20XX_양식문서",
        help="생성된 파일이 저장될 폴더입니다.",
        # label_visibility="collapsed",
    )
    if output_folder:
        st.session_state["output_folder"] = output_folder

    # HWP 양식 및 Excel 데이터 업로드
    main_col1, main_col2 = st.columns(2)
    with main_col1:
        st.header("HWP 양식")
        uploaded_file_hwp = st.file_uploader("HWP", type=["hwp"], label_visibility="collapsed")
    with main_col2:
        st.header("Excel 데이터")
        uploaded_file_excel = st.file_uploader("Excel", type=["xlsx"], label_visibility="collapsed")

    # HWP 양식 열기 및 필드명 가져오기 (Session State)
    if "hwp_field_names" not in st.session_state and uploaded_file_hwp is not None:
        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".hwp",
            dir=output_folder,
        ) as tmp:
            tmp.write(uploaded_file_hwp.read())
            print("Temporary file created:", tmp.name)
            hwp_temp_path = tmp.name

        with hwp_context(visible=False) as hwp_ctx:
            hwp_field_names = open_template(hwp_ctx, hwp_temp_path)

        st.session_state["hwp_temp_path"] = hwp_temp_path
        st.session_state["hwp_field_names"] = hwp_field_names

        st.info(f"양식 문서에서 {len(hwp_field_names)}개의 필드를 찾았습니다.")

    # HWP 양식 필드명 표시 (Session State)
    hwp_field_names = st.session_state.get("hwp_field_names")
    hwp_temp_path = st.session_state.get("hwp_temp_path")
    if hwp_field_names is not None and hwp_temp_path is not None:
        with st.expander("양식문서에 포함된 필드"):
            st.markdown(f"양식문서 임시 저장경로: {hwp_temp_path}")
            st.markdown(f"필드 수: {len(hwp_field_names)}개")
            for field in hwp_field_names:
                st.markdown(f"- {field}")

    # Excel 읽기 및 데이터프레임 생성 (Session State)
    if "df" not in st.session_state and uploaded_file_excel is not None:
        st.session_state["df"] = pd.read_excel(uploaded_file_excel)
    df = st.session_state.get("df")

    # 양식문서 필드 매칭 및 엑셀 데이터 속성 설정
    if df is not None:
        # 열 타입 자동 유추
        inferred_types = {}
        for df_col in df.columns:
            dtype = str(df[df_col].dtype)
            if "datetime" in dtype:
                inferred_types[df_col] = "날짜"
            elif "float" in dtype or "int" in dtype:
                inferred_types[df_col] = "숫자"
            else:
                inferred_types[df_col] = "문자열"

        column_settings = {}
        with st.expander("필드 속성 설정"):
            # 초기화 버튼
            button_col, reset_col = st.columns([1, 1])
            with button_col:
                # 도움말 표시 (modal)
                with st.popover("필드 속성 설정 도움말", use_container_width=True):
                    st.markdown(
                        """
                        `컬럼 (Excel)` Excel 파일의 컬럼명입니다.\n
                        `필드 (HWP)` 값이 입력될 HWP 양식의 필드명입니다.\n
                        `파일명에 포함` 해당 컬럼 값이 파일명에 포함되는지 여부입니다.\n
                        `데이터 타입` 해당 컬럼의 데이터 타입을 설정합니다.\n
                        `포맷` 해당 컬럼의 데이터 포맷을 설정합니다.\n
                        `삭제` 해당 컬럼을 삭제합니다.\n
                        """,
                        unsafe_allow_html=True,
                    )
            with reset_col:
                if st.button(
                    "원래대로 초기화",
                    type="secondary",
                    key="reset_button",
                    use_container_width=True,
                ):
                    if "df" in st.session_state:
                        del st.session_state["df"]
                    if "converted_df" in st.session_state:
                        del st.session_state["converted_df"]
                    st.rerun()

            # 필드 속성 설정
            grid_columns = [1, 1.3, 0.4, 0.7, 1, 0.5]
            grid = st.columns(grid_columns, vertical_alignment="center")
            grid[0].write(
                """
                <div style='text-align: right; height: 0;
                    display: flex; align-items: center; justify-content: right;'>
                    <span style='font-weight: 600;'>컬럼 (Excel)</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
            grid[1].write("**필드 (HWP)**")
            grid[2].write("**파일명**")
            grid[3].write("**데이터 타입**")
            grid[4].write("**포맷**")
            grid[5].write("**삭제**")

            for col in df.columns:
                grid_stack = st.columns(grid_columns, vertical_alignment="center")
                with grid_stack[0]:
                    # 컬럼명 표시 (오른쪽 정렬)
                    st.write(
                        f"""
                        <div style='text-align: right; height: 0;
                            display: flex; align-items: center; justify-content: right;'>
                            <span style='font-weight: 600; font-size: 85%'>{col}</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with grid_stack[1]:
                    field = st.selectbox(
                        f"{col}",
                        options=["-"] + hwp_field_names if hwp_field_names else [],
                        label_visibility="collapsed",
                        key=f"{col}_field",
                    )
                with grid_stack[2]:
                    # 첫번째 열은 기본키로 설정
                    is_pk = st.toggle(
                        "기본키",
                        value=True if col == df.columns[0] else False,
                        label_visibility="collapsed",
                        key=f"{col}_is_pk",
                    )
                with grid_stack[3]:
                    dtype = st.selectbox(
                        "데이터 타입",
                        options=["문자열", "숫자", "날짜"],
                        index=["문자열", "숫자", "날짜"].index(inferred_types[col]),
                        label_visibility="collapsed",
                        key=f"{col}_dtype",
                    )
                with grid_stack[4]:
                    if dtype == "날짜":
                        fmt = st.selectbox(
                            "포맷",
                            options=[
                                "%Y.%m.%d.",
                                "%Y년 %#m월 %#d일",
                                "%Y-%m-%d",
                                "%Y/%m/%d",
                                "%y%m%d",
                            ],
                            label_visibility="collapsed",
                            key=f"{col}_format",
                        )
                    elif dtype == "숫자":
                        fmt = st.selectbox(
                            "포맷",
                            options=["#,##0", "0", "0.00", "0.##"],
                            label_visibility="collapsed",
                            key=f"{col}_format",
                        )
                    else:
                        fmt = st.selectbox(
                            "포맷",
                            options=["(없음)"],
                            label_visibility="collapsed",
                            key=f"{col}_format",
                        )
                with grid_stack[5]:
                    if st.button(
                        "❌",
                        type="tertiary",
                        key=f"{col}_delete",
                    ):
                        df.drop(columns=[col], inplace=True)
                        st.session_state["df"] = df
                        st.session_state["converted_df"] = df.copy()
                        st.rerun()

                column_settings[col] = {
                    "field": field,
                    "is_pk": is_pk,
                    "dtype": dtype,
                    "format": fmt,
                }

        # 기본키로 선택된 열 기준으로 비어있는 행 제거
        _pk_columns = [col for col, setting in column_settings.items() if setting["is_pk"]]
        if _pk_columns:
            df = df.dropna(subset=_pk_columns)
            for col in _pk_columns:
                df = df[df[col].astype(str).str.strip() != ""]

        # 선택된 데이터 타입 및 포맷을 반영
        converted_df = df.copy()
        st.session_state["converted_df"] = converted_df

        # 포맷 및 타입 변환 후 저장
        for col, setting in column_settings.items():
            dtype = setting["dtype"]
            fmt = setting["format"]

            if dtype == "날짜":
                try:
                    converted_df[col] = pd.to_datetime(converted_df[col]).dt.strftime(fmt)
                except Exception:
                    pass  # 변환 실패 시 원본 유지
            elif dtype == "숫자":
                try:
                    converted_df[col] = pd.to_numeric(converted_df[col])
                    if fmt == "#,##0":
                        converted_df[col] = converted_df[col].map("{:,.0f}".format)
                    elif fmt == "0":
                        converted_df[col] = converted_df[col].map("{:.0f}".format)
                    elif fmt == "0.00":
                        converted_df[col] = converted_df[col].map("{:.2f}".format)
                    elif fmt == "0.##":
                        converted_df[col] = converted_df[col].map(
                            lambda x: f"{x:.2f}".rstrip("0").rstrip(".")
                        )
                except Exception:
                    pass

        # 데이터에디터 표시
        st.data_editor(st.session_state["converted_df"], use_container_width=True)

    # 문서 생성 실행 버튼
    if st.button("문서 생성", type="primary", use_container_width=True):
        if uploaded_file_hwp is None:
            st.warning("❗ 먼저 HWP 파일을 업로드하세요.")
        if uploaded_file_excel is None:
            st.warning("❗ 먼저 엑셀 파일을 업로드하세요.")
        if uploaded_file_hwp is not None and uploaded_file_excel is not None:
            with st.spinner("문서 생성 중..."):
                process_documents(
                    dataframe=st.session_state["converted_df"],
                    template_path=st.session_state["hwp_temp_path"],
                    workflow_name=st.session_state["workflow_name"],
                    output_folder=st.session_state["output_folder"],
                    key_columns=[
                        col for col, setting in column_settings.items() if setting["is_pk"]
                    ],
                    # (column_settings[col]["field"], col) -> dict()
                    field_mapping={
                        column_settings[col]["field"]: col
                        for col in st.session_state["converted_df"].columns
                    },
                )
            st.success("문서 생성 완료!")


if __name__ == "__main__":
    main()
