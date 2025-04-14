import winreg
from contextlib import contextmanager
from pathlib import Path
from typing import Dict, List

from pandas import DataFrame
from pyhwpx import Hwp

from hwp.export import save_document
from hwp.template import open_template
from hwp.writer import write_fields


@contextmanager
def hwp_context(visible: bool = False):
    """
    HWP 객체를 생성하고 컨텍스트 매니저로 반환합니다.

    Args:
        visible (bool): HWP 창을 표시할지 여부

    Returns:
        Hwp: HWP 객체
    """
    hwp = Hwp(new=True, visible=visible, register_module=False)
    print("Hwp version:", hwp.Version)

    register_security_module(hwp, "../resource/FilePathCheckerModule.dll")

    try:
        yield hwp
    finally:
        hwp.clear()
        hwp.quit()


def register_security_module(
    hwp: Hwp,
    dll_filepath: str = "FilePathCheckerModule.dll",
) -> bool:
    """
    한글 HwpAutomation에 보안모듈 DLL을 레지스트리에 등록하고, HWP 객체에 등록합니다.
    레지스트리 등록을 위한 관리자 권한이 필요합니다.
    레지스트리에 등록된 DLL은 HWP 객체에 보안모듈로 등록됩니다.

    레지스트리 값 이름을 'FilePathCheckerModule' 이외 이름으로 레지스트리에 등록할 경우
    pyhwpx 패키지 내부에서 `FileNotFoundError` 문제가 발생하여,
    DLL 이름은 'FilePathCheckerModule.dll'로 고정합니다.
        https://github.com/martiniifun/pyhwpx/issues/8

    Args:
        hwp (Hwp): pyhwpx로 생성된 Hwp 객체
        dll_filepath (str): 등록할 보안모듈 DLL 파일 경로
            (상대경로일 경우, 현재 작업 디렉토리 기준으로 해석됨)

    Returns:
        bool: 보안모듈 등록 결과
    """

    dll_path = Path(dll_filepath).resolve()
    if not dll_path.exists():
        raise FileNotFoundError(f"DLL 파일이 존재하지 않습니다: {dll_path}")

    module_name = dll_path.stem  # 'FilePathCheckerModule'
    reg_path = r"Software\HNC\HwpAutomation\Modules"

    # 1-1. 레지스트리 문자열 값(FilePathCheckerModule) 존재 확인
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
            try:
                _, _ = winreg.QueryValueEx(key, module_name)
                # print("[✔] 레지스트리에 보안모듈 등록됨")
            except FileNotFoundError:
                print("[✖] 레지스트리에 보안모듈이 등록되어 있지 않습니다. 새로 등록합니다.")

                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
                    winreg.SetValueEx(key, module_name, 0, winreg.REG_SZ, str(dll_path))
                    print(f"[✔] 레지스트리에 보안모듈 등록 완료: {module_name} → {dll_path}")
    except PermissionError:
        print("[✖] 레지스트리 접근 권한이 없습니다. 관리자 권한으로 실행하세요.")
        return False

    # self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    # self.hwp.XHwpWindows.Active_XHwpWindow.Visible = visible

    # 2. HWP 객체에 보안모듈 등록 (pyhwpx 활용)
    result = hwp.RegisterModule("FilePathCheckDLL", module_name)

    if not result:
        print(f"[✖] HWP 객체에 보안모듈 활성화 실패: {module_name}")
        print(
            rf"""
            아래 링크를 참조하여 보안모듈을 레지스트리에 수동으로 등록하세요. (관리자 권한 필요)
                https://developer.hancom.com/hwpautomation

            DLL 파일을 'C:\\Program Files (x86)\\HNC\\HwpAutomation\\Modules' 폴더(또는 원하는 위치)에 복사한 후,
            레지스트리 편집기를 열고 아래 경로로 이동하여 DLL 경로를 문자열 값으로 등록하세요. (새로 만들기 → 문자열 값)
                컴퓨터\HKEY_CURRENT_USER\Software\HNC\HwpAutomation\Modules
                - 값 이름: {module_name}
                - 값 데이터: {dll_path}
            """
        )

    return result


def process_documents(
    template_path: str,
    dataframe: DataFrame,
    output_folder: str,
    workflow_name: str,
    filename_suffixes: List[str] | None = None,
    key_columns: List[str] | None = None,
    field_mapping: Dict[str, str] | None = None,
) -> None:
    """
    HWP 양식문서 작성(채워넣기) 및 저장

    Args:
        template_path (str): 템플릿 파일 경로
        dataframe (DataFrame): 데이터프레임 객체
        output_folder (str): 출력 폴더 이름
        workflow_name (str): 워크플로우 이름
        filename_suffixes (List[str] | None): 저장할 파일 이름 접미사
            - None인 경우, 기본적으로 인덱스 번호가 사용됨
            - 예: ['001', '002', '003']
        key_columns (List[str] | None): 기본 열 이름들
            - 문서 저장 시 파일 이름에 사용
            - 중복 방지 및 구분을 위해 사용됨
            - 예: ['이름', '생년월일']
        field_mapping (Dict[str, str] | None): 필드 매핑 정보
            - None인 경우, 데이터프레임의 열 이름을 그대로 사용하여 매핑됨
            - 예: {'이름': '이름', '생년월일': '생년월일'}

    Returns:
        None
    """

    if field_mapping is None:
        field_mapping = {col: col for col in dataframe.columns}

    with hwp_context(visible=False) as hwp:
        # 템플릿 열기
        _ = open_template(hwp, template_path)

        # 문서 생성 및 저장
        for index, row in dataframe.iterrows():
            idx = int(str(index)) + 1
            print(f"문서 만드는중... ({idx}/{len(dataframe)})")

            write_fields(hwp, row, field_mapping)

            if filename_suffixes is not None:
                save_filename = f"{workflow_name}_{filename_suffixes[idx - 1]}"
            elif key_columns is not None:
                name_combined = "_".join(str(row[col]) for col in key_columns)
                save_filename = f"{workflow_name}_{name_combined}"
            else:
                save_filename = f"{workflow_name}_{idx}"

            save_document(hwp, output_folder, save_filename)
