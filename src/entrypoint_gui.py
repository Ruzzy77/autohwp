import base64
import os
import tempfile
from difflib import get_close_matches
from pathlib import Path

import pandas as pd
import streamlit as st
from pyjosa.josa import Josa

# pyhwpx가 설치되어 있으면 import
# 윈도우 이외 OS 동작오류 방지
try:
    from hwp.service import hwp_context, open_template, process_documents
except ImportError:
    pass

# UserWarning: Data Validation extension is not supported 해결
# 경고 억제 코드 추가
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def main():
    st.set_page_config(page_title="AutoHWP", layout="wide")
    st.markdown(
        """
        <style>
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
        html, body, [class*="css"]  { font-family: 'Pretendard', sans-serif !important; }
        .st-ax { font-family: 'Pretendard', sans-serif !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <style>
        .st-ay { font-size: 95%; }
        .stSelectbox > div { font-size: 90%; }
        .stMultiSelect [data-baseweb="select"] span { font-size: 90%; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("양식문서 자동완성")
    if "workflow_name" not in st.session_state:
        st.session_state["workflow_name"] = "default_workflow"

    # 워크플로우 이름
    workflow_name = st.text_input(
        "워크플로우 이름",
        # value=st.session_state.get("workflow_name", ""),
        placeholder="예: 20XX_양식문서",
        on_change=lambda: st.session_state.update({"workflow_name": st.session_state.get("workflow_name", "")}),
        max_chars=50,
        help="워크플로우 이름은 생성된 파일명의 접두사로 사용됩니다. (예: 20XX_양식문서_홍길동.hwp)",
    )
    # st.session_state["workflow_name"] = workflow_name

    output_folder_base = Path(tempfile.gettempdir()) / "autohwp"
    if not output_folder_base.exists():
        output_folder_base.mkdir(parents=False, exist_ok=True)

    output_folder = output_folder_base / workflow_name
    # output_folder = st.text_input(
    #     "로컬 저장 폴더",
    #     value=output_folder_base / workflow_name,
    #     placeholder="예: /tmp/20XX_양식문서",
    #     help="임시 파일과 생성된 파일이 저장될 폴더입니다.",
    # )
    st.session_state["output_folder"] = output_folder

    # HWP 양식 및 Excel 데이터 업로드
    main_col1, main_col2 = st.columns(2)
    with main_col1:
        # HWP 양식 헤더와 초기화 버튼
        header_col1, button_col1 = st.columns([2, 1], vertical_alignment="bottom")
        with header_col1:
            st.header("HWP 양식")
        with button_col1:
            if st.button(
                "초기화 및 삭제",
                help="HWP 필드 데이터 초기화 및 임시파일 삭제",
                use_container_width=True,
                type="tertiary",
            ):
                if "hwp_field_names" in st.session_state:
                    del st.session_state["hwp_field_names"]
                if "hwp_file_name" in st.session_state:
                    del st.session_state["hwp_file_name"]
                if "hwp_temp_path" in st.session_state:
                    # HWP 양식 임시 파일 삭제
                    Path(st.session_state["hwp_temp_path"]).unlink(missing_ok=True)
                    del st.session_state["hwp_temp_path"]
                st.session_state["hwp_file_key"] += 1
                st.rerun()

        # HWP 양식 파일 업로드
        if "hwp_file_key" not in st.session_state:
            st.session_state["hwp_file_key"] = 0
        uploaded_file_hwp = st.file_uploader(
            "HWP",
            type=["hwp"],
            label_visibility="collapsed",
            key=f"hwp_file_uploader_{st.session_state['hwp_file_key']}",
        )

    with main_col2:
        # Excel 데이터 헤더와 초기화 버튼
        header_col2, button_col2 = st.columns([5, 1], vertical_alignment="bottom")
        with header_col2:
            st.header("Excel 데이터")
        with button_col2:
            if st.button(
                "초기화",
                help="Excel 데이터 초기화",
                use_container_width=True,
                type="tertiary",
            ):
                if "df" in st.session_state:
                    del st.session_state["df"]
                if "converted_df" in st.session_state:
                    del st.session_state["converted_df"]
                if "excel_file_name" in st.session_state:
                    del st.session_state["excel_file_name"]
                st.session_state["excel_file_key"] += 1
                st.rerun()

        # Excel 파일 업로드
        if "excel_file_key" not in st.session_state:
            st.session_state["excel_file_key"] = 0
        uploaded_file_excel = st.file_uploader(
            "Excel",
            type=["xlsx"],
            label_visibility="collapsed",
            key=f"excel_file_uploader_{st.session_state['excel_file_key']}",
        )

    # HWP 양식 열기 및 필드명 가져오기 (Session State)
    if "hwp_field_names" not in st.session_state and uploaded_file_hwp is not None:
        # HWP 양식 파일을 임시 폴더에 저장
        if "output_folder" in st.session_state:
            output_folder = st.session_state["output_folder"]
            if not Path(output_folder).exists():
                Path(output_folder).mkdir(parents=False, exist_ok=True)

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".hwp",
            dir=output_folder,
            prefix=f"{workflow_name}_upload_template_",
        ) as tmp:
            tmp.write(uploaded_file_hwp.read())
            print("Temporary file created:", tmp.name)
            hwp_temp_path = tmp.name

        with hwp_context(visible=False) as hwp_ctx:
            hwp_field_names = open_template(hwp_ctx, hwp_temp_path)

        st.session_state["hwp_temp_path"] = hwp_temp_path
        st.session_state["hwp_field_names"] = hwp_field_names
        st.session_state["hwp_file_name"] = uploaded_file_hwp.name
        if not workflow_name:
            workflow_name = uploaded_file_hwp.name.split(".")[0]
            st.session_state["workflow_name"] = workflow_name

    # HWP 양식 필드명 표시 (Session State)
    hwp_field_names = st.session_state.get("hwp_field_names")
    hwp_temp_path = st.session_state.get("hwp_temp_path")
    if hwp_field_names is not None and hwp_temp_path is not None:
        with st.expander("양식문서 정보" + f" `{st.session_state['hwp_file_name']}`"):
            st.markdown(
                "<span style='font-size: 85%;'>임시 저장경로</span>" + f" `{hwp_temp_path}`",
                unsafe_allow_html=True,
            )
            st.pills(
                label="필드 목록" + f" `{len(hwp_field_names)}`",
                options=hwp_field_names,
                selection_mode="single",
            )

    # Excel 읽기 및 데이터프레임 생성 (Session State)
    if "df" not in st.session_state and uploaded_file_excel is not None:
        st.session_state["df"] = pd.read_excel(uploaded_file_excel)
        st.session_state["excel_file_name"] = uploaded_file_excel.name

    # 엑셀 데이터 속성 설정
    df = st.session_state.get("df")
    if df is not None:
        # 열 타입 자동 유추
        inferred_types = {}
        for df_col in df.columns:
            df_dtype = str(df[df_col].dtype)
            if "datetime" in df_dtype:
                inferred_types[df_col] = "날짜"
            elif "float" in df_dtype or "int" in df_dtype:
                inferred_types[df_col] = "숫자"
            else:
                inferred_types[df_col] = "문자열"

        column_settings = {}
        with st.expander("데이터 속성 설정" + f" `{st.session_state['excel_file_name']}`"):
            st.markdown(
                "<span style='font-size: 85%;'>컬럼 개수</span>" + f" `{len(df.columns)}`",
                unsafe_allow_html=True,
            )
            st.markdown(
                "<span style='font-size: 85%;'>원본 데이터 개수</span>" + f" `{len(df)}`",
                unsafe_allow_html=True,
            )

            # 초기화 버튼
            button_col, reset_col = st.columns([1, 1])
            with button_col:
                # 도움말 표시 (modal)
                with st.popover("데이터 속성 설정 도움말", use_container_width=True):
                    st.markdown(
                        """
                        `컬럼 (Excel)` 양식에 입력될 값을 가지는 Excel 컬럼명입니다.\n
                        `파일명` 해당 컬럼 값이 파일명에 포함되는지 여부입니다.\n
                        `데이터 타입` 해당 컬럼의 데이터 타입을 설정합니다.\n
                        `포맷` 해당 컬럼의 데이터 포맷을 설정합니다.\n
                        `제거` 해당 컬럼을 변환할 데이터에서 제거합니다.\n
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
            grid_columns = [1, 0.4, 0.7, 1, 0.5]
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
            grid[1].write("**파일명**")
            grid[2].write("**데이터 타입**")
            grid[3].write("**포맷**")
            grid[4].write("**제거**")

            for col in df.columns:
                grid_stack = st.columns(grid_columns, vertical_alignment="center")
                with grid_stack[0]:
                    # 컬럼명 표시 (오른쪽 정렬)
                    st.write(
                        f"""
                        <div style='text-align: right; height: 0.3rem;
                            display: flex; align-items: center; justify-content: right;'>
                            <span style='font-weight: 600; font-size: 100%'>{col}</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with grid_stack[1]:
                    # 첫번째 열은 기본키로 설정
                    _is_key = st.toggle(
                        "기본키",
                        value=True if col == df.columns[0] else False,
                        label_visibility="collapsed",
                        key=f"{col}_is_key",
                    )
                with grid_stack[2]:
                    _dtype = st.selectbox(
                        "데이터 타입",
                        options=["문자열", "숫자", "날짜"],
                        index=["문자열", "숫자", "날짜"].index(inferred_types[col]),
                        label_visibility="collapsed",
                        key=f"{col}_dtype",
                    )
                with grid_stack[3]:
                    if _dtype == "날짜":
                        _fmt = st.selectbox(
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
                    elif _dtype == "숫자":
                        _fmt = st.selectbox(
                            "포맷",
                            options=["#,##0", "0", "0.00", "0.##"],
                            label_visibility="collapsed",
                            key=f"{col}_format",
                        )
                    else:
                        _fmt = st.selectbox(
                            "포맷",
                            options=[],
                            label_visibility="collapsed",
                            key=f"{col}_format",
                        )
                with grid_stack[4]:
                    if _is_removed := st.button(
                        "❌",
                        type="tertiary",
                        key=f"{col}_isremove",
                    ):
                        df.drop(columns=[col], inplace=True)
                        st.session_state["df"] = df
                        st.session_state["converted_df"] = df.copy()
                        st.rerun()

                column_settings[col] = {
                    "is_key": _is_key,
                    "dtype": _dtype,
                    "format": _fmt,
                    "is_removed": _is_removed,
                }

            # 기본키로 선택된 열 기준으로 비어있는 행 제거
            _pk_columns = [col for col, setting in column_settings.items() if setting["is_key"]]
            if _pk_columns:
                df = df.dropna(subset=_pk_columns)
                for col in _pk_columns:
                    df = df[df[col].astype(str).str.strip() != ""]

            # 기본키에 해당하는 값들을 "_"로 합쳐서 파일명으로 사용
            if len(_pk_columns) > 0:
                st.session_state["save_filenames"] = (
                    df[_pk_columns].astype(str).agg("_".join, axis=1)
                )
            else:
                st.session_state["save_filenames"] = df.index.astype(str).tolist()

            # 선택된 데이터 타입 및 포맷을 반영
            converted_df = df.copy()
            st.session_state["converted_df"] = converted_df

            # 포맷 및 타입 변환
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
            st.divider()
            st.markdown(
                "<span style='font-size: 85%;'>필터링 데이터 개수</span>"
                + f" `{len(converted_df)}`",
                unsafe_allow_html=True,
            )
            st.data_editor(st.session_state["converted_df"], use_container_width=True)

    # 필드 매칭 및 추가 전처리
    if (
        st.session_state.get("hwp_field_names") is not None
        and st.session_state.get("converted_df") is not None
    ):
        field_settings = {}
        with st.expander("필드 매칭"):
            # | 필드 | 컬럼 (Excel) | 사용자 지정값 |
            # | 텍스트 | 멀티셀렉트 | 텍스트 입력 |

            grid_match_columns = [0.8, 2.2, 1]
            grid_match = st.columns(grid_match_columns, vertical_alignment="center")

            grid_match[0].markdown(
                f"""
                <div style='text-align: right; height: 0;
                    display: flex; align-items: center; justify-content: right;'>
                    <span style='color: gray; font-size: 70%;'>{len(st.session_state["hwp_field_names"])}</span>
                    &nbsp;&nbsp;&nbsp;
                    <span style='font-weight: 600;'>필드</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
            grid_match[1].write(
                "**컬럼 (Excel)**"
                "&nbsp;&nbsp;"
                "<span style='color: gray; font-size: 70%;'>다중선택 가능 (3개)</span>",
                unsafe_allow_html=True,
            )
            grid_match[2].write("**사용자 지정 값**")

            for field in st.session_state["hwp_field_names"]:
                matching_stack = st.columns(grid_match_columns, vertical_alignment="center")
                with matching_stack[0]:
                    # 필드명 표시 (오른쪽 정렬)
                    st.write(
                        f"""
                        <div style='text-align: right; height: 0.0rem;
                            display: flex; align-items: center; justify-content: right;'>
                            <span style='font-weight: 600; font-size: 80%'>{field}</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with matching_stack[1]:
                    _selected_columns = st.multiselect(
                        "컬럼 (Excel)",
                        options=["사용자 지정", "조사(은/는)"] + list(converted_df.columns),
                        placeholder="하나 이상 선택하세요.",
                        max_selections=3,
                        default=get_close_matches(field, converted_df.columns, n=1, cutoff=0.4),
                        label_visibility="collapsed",
                        key=f"{field}_excel_columns",
                    )
                with matching_stack[2]:
                    _option_selected = any(
                        c in _selected_columns for c in ["사용자 지정", "조사(은/는)"]
                    )
                    _fixed_value = st.text_input(
                        "사용자 지정",
                        value="",
                        label_visibility="collapsed",
                        placeholder="값을 입력하세요." if _option_selected else "",
                        disabled=not _option_selected,
                        key=f"{field}_fixed_value",
                    )

                field_settings[field] = {
                    "columns": _selected_columns,
                    "fixed_value": _fixed_value,
                }

        # 필드명으로 컬럼 구성된 데이터프레임 뷰어
        # "사용자 지정"은 컬럼 이름이 아니라, 해당 컬럼에 들어갈 값을 입력함
        field_df = pd.DataFrame(columns=hwp_field_names)
        for field, setting in field_settings.items():
            selected_columns = setting["columns"]
            fixed_value = setting["fixed_value"]

            # KeyError: "['사용자 지정'] not in index" 해결
            # 선택된 열이 데이터프레임에 존재하는지 확인하는 검증 로직 추가
            if "converted_df" in st.session_state:
                selected_columns = [col for col in st.session_state["converted_df"].columns if col in converted_df.columns]
                if not selected_columns:
                    st.error("선택한 열이 데이터프레임에 존재하지 않습니다.")
                    return

            if len(selected_columns) == 0:
                st.warning(f"❗ 필드 `{field}`에 대해 선택된 컬럼이 없습니다. 필드를 선택하세요.")
                continue
            elif len(selected_columns) == 1:
                # 단일 선택 컬럼이므로 셀 데이터 그대로 넣기
                if selected_columns[0] == "사용자 지정":
                    # 사용자 지정값이므로 고정값 사용
                    field_df[field] = fixed_value
                elif selected_columns[0] == "조사(은/는)":
                    # 단독사용 불가
                    st.warning("❗ 조사(은/는) 단독 사용이 지원되지 않습니다.")
                else:
                    field_df[field] = converted_df[selected_columns[0]]
            elif len(selected_columns) > 1:
                # 다중 선택 컬럼이므로 셀 데이터 합치기
                if "사용자 지정" in selected_columns and len(selected_columns) == 3:
                    # "사용자 지정"이 포함된 경우 (3개 선택: df 컬럼 값 + 사용자 지정 값 + df 컬럼 값)
                    selected_columns.remove("사용자 지정")
                    field_df[field] = converted_df[selected_columns].apply(
                        lambda row: f" {fixed_value} ".join(row.values.astype(str)), axis=1
                    )
                elif "조사(은/는)" in selected_columns and len(selected_columns) == 2:
                    # "조사(은/는)"이 포함된 경우 (2개 선택: df 컬럼 값 + 조사(은/는))
                    selected_columns.remove("조사(은/는)")
                    try:
                        fixed_value = st.text_input("사용자 지정", value="", label_visibility="collapsed")
                        if not fixed_value:
                            raise ValueError("올바르지 않은 조사 값입니다.")
                        field_df[field] = converted_df[selected_columns].apply(
                            lambda row: Josa.get_full_string(row.values.astype(str)[0], fixed_value),
                            axis=1,
                        )
                    except ValueError as e:
                        st.error(str(e))
                else:
                    # 실제 컬럼으로만 구성된 경우 (df 컬럼 값 + df 컬럼 값 + df 컬럼 값)
                    field_df[field] = converted_df[selected_columns].apply(
                        lambda row: " ".join(row.values.astype(str)), axis=1
                    )

        st.session_state["field_df"] = field_df
        st.session_state["field_settings"] = field_settings
        st.session_state["column_settings"] = column_settings

        try:
            st.dataframe(st.session_state["field_df"], use_container_width=True)
        except Exception as e:
            st.error(f"데이터프레임 표시 중 오류 발생: {e}")

    st.divider()

    # 문서 생성 실행 버튼
    button_col1, button_col2 = st.columns([1, 1])
    with button_col1:
        if st.button("문서 생성", type="primary", use_container_width=True):
            if uploaded_file_hwp is None:
                st.warning("❗ 먼저 HWP 파일을 업로드하세요.")
            if uploaded_file_excel is None:
                st.warning("❗ 먼저 엑셀 파일을 업로드하세요.")
            if uploaded_file_hwp is not None and uploaded_file_excel is not None:
                with st.spinner("문서 생성 중..."):
                    process_documents(
                        dataframe=st.session_state["field_df"],
                        template_path=st.session_state["hwp_temp_path"],
                        workflow_name=st.session_state["workflow_name"],
                        output_folder=st.session_state["output_folder"],
                        filename_suffixes=st.session_state["save_filenames"],
                    )
                st.success("문서 생성 완료!")

    # 문서 삭제 버튼 (hwp, pdf)
    with button_col2:
        if st.button("모두 삭제", type="secondary", use_container_width=True):
            if "output_folder" in st.session_state:
                output_folder = st.session_state["output_folder"]
                if Path(output_folder).exists():
                    # output_folder에 생성된 파일 목록 삭제
                    # workflow_name_로 시작하는 파일 삭제
                    for file in Path(output_folder).glob(f"{workflow_name}_*"):
                        os.remove(file)
                    st.success("모든 파일 삭제 완료!")
                    # st.rerun()
                else:
                    st.warning("삭제할 파일이 없습니다.")

    # 생성문서 파일 목록 뷰어
    # output_folder에 생성된 파일 목록 표시 (hwp, pdf)
    # flexbox에서 가로로 스택되며,
    # hwp / pdf 아이콘으로 표시하고, 아이콘을 누르면 다운로드
    # 아이콘 아래 파일명 표시
    if "output_folder" in st.session_state:
        output_folder = st.session_state["output_folder"]
        if not Path(output_folder).exists():
            st.header("시작하려면 HWP 양식과 Excel 파일을 업로드하세요.")
        else:
            patterns = ["*.hwp", "*.pdf"]
            file_paths = [
                file for pattern in patterns for file in Path(output_folder).glob(pattern)
            ]

            st.header("생성된 문서")
            if not file_paths:
                st.write(
                    "<span style='color: gray; font-size: 14px'> 생성된 문서가 없습니다.</span>",
                    unsafe_allow_html=True,
                )
            else:
                st.write(
                    """<span style='color: gray; font-size: 14px'>
                        아이콘을 클릭해 다운로드합니다.
                        <br>파일명은 워크플로우 이름을 기준으로 생성됩니다. (예: 20XX_양식문서_홍길동.hwp)
                    </span>""",
                    unsafe_allow_html=True,
                )
            st.markdown(
                """
                <style>
                    .file-container {
                        display: flex;
                        flex-wrap: wrap;
                        gap: 5px;
                    }
                    .file-item {
                        display: flex;
                        flex-direction: column;
                        align-items: center;
                        width: 120px;
                    }
                    .file-item img {
                        width: auto;
                        height: auto;
                    }
                    .file-name {
                        font-size: 14px;
                        text-align: center;
                        word-wrap: break-word;
                        overflow-wrap: break-word;
                        width: 100%;
                    }
                </style>
                """,
                unsafe_allow_html=True,
            )

            file_items = ""
            for file_path in file_paths:
                file_name = file_path.name
                file_type = file_path.suffix.lower()
                mime_type = "application/pdf" if file_type == ".pdf" else "application/x-hwp"

                file_path_b64 = (
                    f"data:{mime_type};base64,{base64.b64encode(file_path.read_bytes()).decode()}"
                )
                file_name_short = file_name.replace(f"{st.session_state['workflow_name']}_", "")

                resource_path = Path(__file__).parent / "icons"
                image_icons = {
                    ".hwp": "hwp_icon.png",
                    ".pdf": "pdf_icon.png",
                    "default": "default_icon.png",
                }
                file_type = image_icons.get(file_type, image_icons["default"])
                icon_path = f"{resource_path}/{file_type}"
                icon_path_b64 = f"data:image/png;base64,{base64.b64encode(Path(icon_path).read_bytes()).decode()}"

                file_items += f'''
                <div class="file-item">
                    <a href="{file_path_b64}" download="{file_name}">
                        <img src="{icon_path_b64}" alt="{file_type} icon">
                    </a>
                    <span class="file-name">{file_name_short}</span>
                </div>
                '''

            # 전체 파일 목록을 Flexbox 컨테이너에 삽입
            st.markdown(
                f"""
                <div class="file-container">
                    {file_items}
                </div>
                """,
                unsafe_allow_html=True,
            )


if __name__ == "__main__":
    main()
