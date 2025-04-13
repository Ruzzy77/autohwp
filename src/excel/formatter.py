from datetime import datetime

from openpyxl.cell import Cell, MergedCell
from openpyxl.styles import is_date_format

from excel.config import EXCEL_TO_STRFTIME


def format_cell(cell: "Cell | MergedCell") -> str | None:
    """
    셀의 서식을 확인해 날짜라면 포맷된 문자열 반환

    Args:
        cell (Cell | MergedCell): 엑셀 셀 객체

    Returns:
        str | None: 포맷된 문자열 또는 None
    """
    if cell.value is None:
        return None

    if is_date_format(cell.number_format) and isinstance(cell.value, datetime):
        fmt = EXCEL_TO_STRFTIME.get(cell.number_format, "%Y.%m.%d.")
        if fmt:
            return cell.value.strftime(fmt)

    return str(cell.value)
