# %% entrypoint_cli.py

from config import FIELD_MAPPING, Config
from excel.loader import data_loader, load_worksheet
from excel.preprocess import preprocess_dataframe
from hwp.service import process_documents

if __name__ == "__main__":
    print("Hello from autohwp!")

    UI_INPUTS = {
        "const_sanhak": "산학협력단장명",
        "sanhak_name": "김응태",
        "workflow_name": "프로젝트연구교수 채용계약서",
        "key_columns": ["성명", "사번"],
    }
    PATH_CONFIG = {
        "template_path": "../template/contract.hwp",
        "excel_path": "../template/contract_fill.xlsx",
    }

    excel_path = PATH_CONFIG["excel_path"]
    ws = load_worksheet(excel_path)
    df = data_loader(ws, key_columns=UI_INPUTS["key_columns"])

    df = preprocess_dataframe(df, UI_INPUTS)
    df = df.dropna(how="all")

    # 임시로 설정값을 전달하여 process_documents 호출
    process_documents(
        template_path=PATH_CONFIG["template_path"],
        dataframe=df,
        output_folder=UI_INPUTS["workflow_name"],
        workflow_name=UI_INPUTS["workflow_name"],
        key_columns=UI_INPUTS["key_columns"],
        field_mapping=FIELD_MAPPING,
    )

    print("Documents generated successfully.")
