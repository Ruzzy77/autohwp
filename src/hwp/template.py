from pyhwpx import Hwp


def open_template(hwp: Hwp, path: str) -> list[str]:
    """
    HWP 템플릿을 열고 필드 목록을 반환합니다.

    Args:
        hwp (Hwp): HWP 객체
        path (str): 템플릿 파일 경로

    Returns:
        list[str]: 필드 이름 목록
    """
    hwp.open(path)
    print("Opened document title:", hwp.get_title())

    fields_hwp = hwp.get_field_list()
    # fields_hwp = "title{{0}}\x02body{{0}}\x02title{{1}}\x02body{{1}}\x02footer{{0}}"
    # 구분자 "\x02"로 필드를 나눔
    # 각 필드의 "{0}" 또는 "{1}"은 인덱스 번호를 나타냄

    # fields = [ "title{{0}}", "body{{0}}", "title{{1}}", "body{{1}}", "footer{{0}}" ]

    # 필드 이름을 추출하여 리스트로 변환
    fields = [field.strip() for field in fields_hwp.split("\x02") if field.strip()]

    print("Total number of fields:", len(fields))

    return fields
