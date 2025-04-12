from typing import Any, Dict

from pydantic import BaseModel, Field

UI_CONFIG = {
    "const_sanhak": "산학협력단장명",
    "sanhak_name": "김응태",
    "workflow_name": "프로젝트연구교수 채용계약서",
}

PATH_CONFIG = {
    "template_path": "template/contract.hwp",
    "excel_path": "template/contract_fill.xlsx",
}

PROCESS_CONFIG = {
    "column_row": 1,
    "start_row": 2,
    "primary_column": "성명",
}

FIELD_MAPPING = {
    "책임교수명(본문){{0}}": "책임교수명(본문)",
    "계약당사자명(본문){{0}}": "성명(본문)",
    "프로젝트명{{0}}": "프로젝트명",
    "총 사업기간{{0}}": "총 사업기간",
    "당해연도 사업기간{{0}}": "당해연도 사업기간",
    "계약시작일{{0}}": "계약시작일",
    "계약종료일{{0}}": "계약종료일",
    "총 계약금액{{0}}": "총 계약금액",
    "월 계약금액{{0}}": "월 계약금액",
    "급여일{{0}}": "급여일",
    "계약일{{0}}": "계약일",
    "산학협력단장명(서명란){{0}}": "산학협력단장명",
    "책임교수명(서명란){{0}}": "책임교수명",
    "주소{{0}}": "주소",
    "계약당사자명(서명란){{0}}": "성명",
    "휴대폰번호{{0}}": "휴대폰번호",
}
