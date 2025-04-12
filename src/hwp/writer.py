import pandas as pd
from pyhwpx import Hwp


def write_fields(hwp: Hwp, row: pd.Series, mapping: dict[str, str]) -> None:
    """
    HWP 필드에 DataFrame 행 데이터를 입력합니다.

    Args:
        hwp (Hwp): HWP 객체
        row (pd.Series): 데이터프레임의 행
        mapping (dict[str, str]): 필드와 컬럼 매핑 딕셔너리

    Returns:
        None
    """
    for field_name, column_name in mapping.items():
        column_value = row[column_name]
        hwp.put_field_text(field_name, column_value)
