from pathlib import Path
from unicodedata import normalize

from pathvalidate import sanitize_filename
from pyhwpx import Hwp


def save_document(hwp: Hwp, folder: str, filename: str) -> None:
    """
    HWP 문서를 저장하고 PDF로 변환합니다.

    Args:
        hwp (Hwp): HWP 객체
        folder (str): 저장 폴더 경로
        filename (str): 저장 파일 이름

    Returns:
        None
    """

    folder = sanitize_filename(folder, platform="Windows", replacement_text="_")
    filename = sanitize_filename(filename, platform="Windows", replacement_text="_")

    # 파일경로 한글 자소분리 문제 해결 (NFD -> NFC)
    folder = normalize("NFC", folder)
    filename = normalize("NFC", filename)

    save_folder_path = (Path("temp") / folder).resolve().absolute()
    if not save_folder_path.exists():
        save_folder_path.mkdir(parents=True, exist_ok=True)

    save_path = save_folder_path / f"{filename}.hwp"

    if hwp.save_as(str(save_path)):
        print(f"HWP 저장 완료: {save_path.name}")

    if hwp.save_as(str(save_path.with_suffix(".pdf")), format="pdf"):
        print(f"PDF 저장 완료: {save_path.with_suffix('.pdf').name}")
