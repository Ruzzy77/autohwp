from typing import Any, Dict

from pydantic import BaseModel, Field

FIELD_MAPPING = {  # 추후 UI 및 JSON으로 입력받을 수 있도록 수정
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


class Config(BaseModel):
    """
    어플리케이션 전반에 걸쳐 사용되는 설정을 정의하는 클래스입니다.
    이 클래스는 각종 설정 값의 유효성을 검사하고, 기본값을 설정하는 데 사용됩니다.
    또한, 설정 값들을 JSON 형식으로 직렬화하는 기능도 제공합니다.
    """

    class Config:
        """
        Pydantic 모델의 설정을 정의하는 클래스입니다.
        이 클래스는 모델의 동작 방식을 조정하는 다양한 설정을 포함합니다.
        """

        arbitrary_types_allowed = True
        json_encoders = {str: lambda v: v}
        validate_by_name = True
        use_enum_values = True
        validate_assignment = True

    workflow_name: str = Field(
        ...,
        title="워크플로우 이름",
        description="워크플로우 결과를 저장하는 폴더의 이름으로 사용됩니다.",
        examples=["Workflow 1", "채용계약서"],
        min_length=1,
        max_length=50,
    )

    template_path: str = Field(
        ...,
        title="HWP 템플릿 경로",
        description="HWP 템플릿 파일의 경로입니다.",
        examples=["template/contract.hwp"],
    )
    excel_path: str = Field(
        ...,
        title="엑셀 템플릿 경로",
        description="엑셀 템플릿 파일의 경로입니다.",
        examples=["template/contract_fill.xlsx"],
    )

    header_row: int = Field(
        ...,
        title="헤더 행 번호",
        description="엑셀에서 열 제목이 있는 행 번호 (1부터 시작)",
        ge=1,
    )
    start_row: int = Field(
        ...,
        title="시작 행 번호",
        description="엑셀에서 데이터가 시작되는 행 번호",
        ge=1,
    )
    end_row: int = Field(
        ...,
        title="끝 행 번호",
        description="엑셀에서 데이터가 끝나는 행 번호",
        ge=1,
    )
    key_columns: list[str] = Field(
        ...,
        title="기본 키 열",
        description="엑셀에서 기본 키가 되는 열의 리스트입니다.",
        examples=[["성명", "사번"]],
        min_length=1,
    )

    field_mapping: Dict[str, str] = Field(
        ...,
        title="필드 매핑 정보",
        description="필드 매핑 정보는 HWP 템플릿에서 사용되는 필드 이름과 엑셀 열 제목을 매핑하는 데 사용됩니다.",
        examples=[
            {
                "책임교수명(본문){{0}}": "책임교수명(본문)",
                "계약당사자명(본문){{0}}": "성명(본문)",
                "프로젝트명{{0}}": "프로젝트명",
            },
        ],
    )
