# AutoHWP

:warning: 현재 개발 중입니다. 완전히 작동하지 않을 수 있습니다.

## 한글(HWP) 문서의 양식 채워넣기 자동화

한글(HWP) 문서의 반복되는 양식을 자동으로 채워넣는 프로그램입니다.

이 프로그램은 한글 문서의 특정 필드를 찾아서 사용자가 제공한 데이터를 자동으로 입력합니다.

이를 통해 똑같은 양식 채워넣기의 반복을 엑셀에 정리한 자료를 통해 한 번에 처리할 수 있습니다.

## 사용법

### :blue_book: 양식 파일 준비하기

양식 파일은 한글(HWP) 문서로 준비합니다.

1. 한글(HWP) 문서를 열고, 양식으로 사용할 부분에 `필드(누름틀)`를 추가합니다.
  > `입력` > `개체` > `필드 입력(Ctrl+K,E)` 또는 `입력` > `누름틀`
2. 필드(누름틀)에 필드 이름을 지정합니다.
  > - 필드 이름은 엑셀 파일의 첫 번째 행(제목 행)과 일치해야 합니다.
  > - 가급적 필드 이름은 중복되지 않도록 설정합니다.

### :green_book: 엑셀 파일 준비하기

엑셀 파일은 양식에 채워넣을 데이터를 준비합니다.
1. 엑셀 파일의 첫 번째 행(제목 행)에 필드 이름을 입력합니다.
2. 두 번째 행부터는 채워넣을 데이터를 입력합니다.
3. 엑셀 파일을 저장합니다.
  > - 엑셀 파일은 `.xlsx` 형식으로 저장합니다.
  > - 엑셀 파일의 첫 번째 시트에 필드 이름과 데이터가 있어야 합니다.

### :gem: 웹 GUI 사용하기

1. 웹 GUI를 통해 준비된 양식 파일과 엑셀 파일을 업로드합니다.
2. 업로드가 완료되면, 데이터 내용과 필드 값 서식을 확인하고 수정합니다.
3. 완료 후, `양식 채우기` 버튼을 클릭합니다.
4. 양식 채우기가 완료되면, `다운로드` 버튼을 클릭하여 결과 파일을 다운로드합니다.

<br/><br/>

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
- pathvalidate `latest`
  > 파일 경로 유효성 검사를 위해 사용합니다.

### Development

- ruff
- notebook

### GUI

- nicegui `latest`
  > 웹 기반 GUI 및 백엔드 서버를 제공합니다.

<br/><br/>

## TODO

- [ ] 웹 기반 (NiceGUI) 구현
- [ ] 개인 배포용 pyinstaller로 exe 파일 만들기
