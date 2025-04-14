from pathlib import Path
from tempfile import gettempdir
from unicodedata import normalize

from pathvalidate import sanitize_filename
from pyhwpx import Hwp


def save_document(hwp: Hwp, folderNameOrPath: str, filename: str) -> None:
    """
    HWP 문서를 저장하고 PDF로 변환합니다.

    Args:
        hwp (Hwp): HWP 객체
        folderNameOrPath (str): 저장 폴더 이름 또는 경로
        filename (str): 저장 파일 이름

    Returns:
        None
    """

    folderpath = Path(folderNameOrPath)
    # 경로로 주어지면 경로에 저장, 폴더이름으로 주어지면 temp 폴더에 저장
    if not folderpath.is_dir():
        folderpath = Path(gettempdir()) / folderpath

    foldername = folderpath.name if folderpath.is_dir() else folderpath.stem

    foldername = sanitize_filename(foldername, platform="Windows", replacement_text="_")
    filename = sanitize_filename(filename, platform="Windows", replacement_text="_")

    # 파일경로 한글 자소분리 문제 해결 (NFD -> NFC)
    foldername = normalize("NFC", foldername)
    filename = normalize("NFC", filename)

    # save_folder_path = (Path("temp") / folder).resolve().absolute()
    save_folder_path = folderpath.parent / foldername
    if not save_folder_path.exists():
        save_folder_path.mkdir(parents=True, exist_ok=True)

    save_path = save_folder_path / f"{filename}.hwp"

    if hwp.save_as(str(save_path)):
        print(f"HWP 저장 완료: {save_path.name}")

    if hwp.save_as(str(save_path.with_suffix(".pdf")), format="pdf"):
        print(f"PDF 저장 완료: {save_path.with_suffix('.pdf').name}")
