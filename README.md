# AutoHWP

:warning: 현재 개발 중입니다. 완전히 작동하지 않을 수 있습니다.


## 한글(HWP) 문서의 양식 채워넣기 자동화

한글(HWP) 문서의 반복되는 양식을 자동으로 채워넣는 프로그램입니다.

이 프로그램은 한글 문서의 특정 필드를 찾아서 사용자가 제공한 데이터를 자동으로 입력합니다.

이를 통해 똑같은 양식 채워넣기의 반복을 엑셀에 정리한 자료를 통해 한 번에 처리할 수 있습니다.


## 사용법

### 양식 파일 준비하기
양식 파일은 한글(HWP) 문서로 준비합니다.
1. 한글(HWP) 문서를 열고, 양식으로 사용할 부분에 `필드(누름틀)`를 추가합니다.
  > `입력` > `개체` > `필드 입력(Ctrl+K,E)` 또는 `입력` > `누름틀`
1. 필드(누름틀)에 필드 이름을 지정합니다.
2. 문서를 저장하고, `template` 폴더에 넣습니다.

### 엑셀 파일 준비하기
엑셀 파일은 양식에 채워넣을 데이터를 준비합니다.
1. 엑셀 파일의 첫 번째 행(제목 행)에 필드 이름을 입력합니다.
2. 두 번째 행부터는 채워넣을 데이터를 입력합니다.
3. 엑셀 파일을 저장하고, `data` 폴더에 넣습니다.

## TODO
- [ ] 웹 기반 (NiceGUI) 구현
- [ ] 개인 배포용 pyinstaller로 exe 파일 만들기


## Python Requirements
> `pyproject.toml` 또는 `requirements.txt` 파일을 통해 패키지를 설치할 수 있습니다.

### Dependencies
- python `>= 3.12`
- pandas `latest`
  > openpyxl을 사용하여 엑셀 파일을 읽고 씁니다.
- pyhwpx `==0.50.32`
  > 최신버전 `0.50.36`이지만, 라이브러리 오류로 인해 `0.50.32`로 고정  
  > - `pywin32`를 내부적으로 포함하므로, Windows에서만 작동합니다.
  > - 한글(HWP) 프로그램이 설치되어 있어야 합니다.
- pyjosa `latest`
  > 한글 이름의 조사 처리를 위해 사용합니다. (홍길동 -> 홍길동이/홍길동을)

### Development
- ruff
- notebook

### GUI
- nicegui `latest`
  > 웹 기반 GUI 및 백엔드 서버를 제공합니다.
