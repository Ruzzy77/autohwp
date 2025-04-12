# %%
from src.config import PATH_CONFIG, PROCESS_CONFIG, UI_CONFIG
from src.hwp.service import process_documents

if __name__ == "__main__":
    print("Hello from autohwp!")

    # 임시로 설정값을 전달하여 process_documents 호출
    process_documents(
        template_path=PATH_CONFIG["template_path"],
        excel_path=PATH_CONFIG["excel_path"],
        output_folder=UI_CONFIG["workflow_name"],
        config={**UI_CONFIG, **PROCESS_CONFIG},
    )

    print("Documents generated successfully.")
