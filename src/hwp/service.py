from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict

from pandas import DataFrame
from pyhwpx import Hwp

from src.excel.formatter import format_cell
from src.excel.loader import load_worksheet
from src.excel.preprocess import preprocess_dataframe
from src.hwp.export import save_document
from src.hwp.field_mapper import FIELD_MAPPING
from src.hwp.template import open_template
from src.hwp.writer import write_fields


@contextmanager
def hwp_context(visible: bool = False):
    """
    HWP 객체를 생성하고 컨텍스트 매니저로 반환합니다.

    Args:
        visible (bool): HWP 창을 표시할지 여부

    Returns:
        Hwp: HWP 객체
    """
    hwp = Hwp(new=True, visible=visible)
    try:
        yield hwp
    finally:
        hwp.clear()
        hwp.quit()


def process_documents(
    template_path: str,
    excel_path: str,
    output_folder: str,
    config: Dict[str, Any],
) -> None:
    """
    HWP 양식문서 작성(채워넣기) 및 저장

    Args:
        template_path (str): 템플릿 파일 경로
        excel_path (str): 엑셀 파일 경로
        output_folder (str): 출력 폴더 이름
        config (Dict[str, str]): 설정 값 딕셔너리

    Returns:
        None
    """
    with hwp_context(visible=False) as hwp:
        print("Hwp version:", hwp.Version)

        # 템플릿 열기
        open_template(hwp, template_path)

        # 엑셀 데이터 로드 및 처리
        ws = load_worksheet(excel_path, config)
        header = [
            cell.value
            for cell in next(
                ws.iter_rows(min_row=config["column_row"], max_row=config["column_row"])
            )
        ]
        primary_column_index = header.index(config["primary_column"])

        fill_data = []
        for row in ws.iter_rows(min_row=config["start_row"]):
            if row[primary_column_index].value is not None:
                fill_data.append([format_cell(cell) for cell in row])

        df = DataFrame(fill_data, columns=header)
        df = preprocess_dataframe(df, config)

        # 문서 생성 및 저장
        for index, row in df.iterrows():
            idx = int(str(index)) + 1
            print(f"문서 만드는중... ({idx}/{len(df)})")

            write_fields(hwp, row, FIELD_MAPPING)

            save_filename = f"{config['workflow_name']}_{row[config['primary_column']]}"
            save_document(hwp, output_folder, save_filename)
