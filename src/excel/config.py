# 엑셀 서식 관련 상수 정의
EXCEL_TO_STRFTIME = {
    r'yyyy"년"\ m"월"\ d"일";@': "%Y년 %#m월 %#d일",  # Windows 포맷
    r"yyyy/mm/dd/": "%Y.%m.%d.",
    # 필요 시 추가 매핑
}
