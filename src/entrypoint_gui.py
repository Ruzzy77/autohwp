# %%
# entrypoint_gui.py
# 문서 자동화 시스템 - Apple 스타일 UI 레이아웃 (기능 없이 UI 구성만)

from nicegui import ui

# ----------------------------
# 공통 스타일 프리셋
# ----------------------------

APPLE_FONT = "font-sans text-gray-800"
SECTION_STYLE = "rounded-xl shadow-md p-6 bg-white"
TITLE_STYLE = "text-3xl font-bold text-gray-900"
SUBTITLE_STYLE = "text-lg font-semibold mb-2 text-gray-700"
BUTTON_STYLE = "bg-black text-white px-4 py-2 rounded-lg hover:bg-gray-800 transition"

# ----------------------------
# UI 레이아웃 구성
# ----------------------------

# 기존: 전체를 하나의 column으로 구성
# 변경 후: 2x2 grid 구성

ui.add_head_html(
    "<style>body { background-color: #f9f9fa; font-family: -apple-system, BlinkMacSystemFont, sans-serif; }</style>"
)
ui.label("📄 문서 자동화 시스템").classes(f"{TITLE_STYLE} mt-6 text-center")

with ui.row().classes("w-full max-w-6xl mx-auto gap-6 mt-8"):
    with ui.column().classes("gap-6 w-full"):
        # 섹션 1: 엑셀 업로드
        with ui.card().classes(SECTION_STYLE):
            ui.label("1. 엑셀 파일 업로드").classes(SUBTITLE_STYLE)
            ui.upload(label="파일 선택 (.xlsx)", auto_upload=True).props("accept=.xlsx")
            ui.label("업로드된 파일이 없습니다.").style("color: #999; font-size: 0.9rem")

        # 섹션 2: 데이터 미리보기
        with ui.card().classes(SECTION_STYLE):
            ui.label("2. 데이터 미리보기").classes(SUBTITLE_STYLE)
            ui.table(columns=[], rows=[], row_key="index").classes("w-full")
            ui.button("데이터 미리보기 실행").classes(BUTTON_STYLE).props("icon=visibility")

    with ui.column().classes("gap-6 w-full"):
        # 섹션 3: 템플릿 및 고정값 설정
        with ui.card().classes(SECTION_STYLE):
            ui.label("3. 템플릿 및 고정값 입력").classes(SUBTITLE_STYLE)
            ui.input("책임교수명").classes("w-full")
            ui.input("계약연도 (예: 2025)").classes("w-full")
            ui.select(["템플릿A", "템플릿B"], label="템플릿 선택").classes("w-full")

        # 섹션 4+5 합침: 문서 생성 및 결과 다운로드
        with ui.card().classes(SECTION_STYLE):
            ui.label("4. 문서 생성 및 다운로드").classes(SUBTITLE_STYLE)
            ui.button("문서 생성 실행").classes(BUTTON_STYLE).props("icon=upload")
            ui.label("생성된 문서를 다운로드할 수 있는 영역입니다.").style(
                "color: #999; font-size: 0.9rem"
            )

ui.run(title="문서 자동화 시스템", port=8080, reload=True)
