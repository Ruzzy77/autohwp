from openpyxl import load_workbook


def load_worksheet(path: str, config: dict):
    """
    엑셀 파일을 로드하고 워크시트를 반환합니다.

    Args:
        path (str): 엑셀 파일 경로
        config (dict): 설정 값

    Returns:
        Worksheet: 엑셀 워크시트 객체
    """
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    if ws is None:
        raise ValueError("No active worksheet found in the Excel file.")
    return ws
