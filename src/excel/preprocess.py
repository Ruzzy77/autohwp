import pandas as pd
from pyjosa.josa import Josa


def preprocess_dataframe(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    """
    DataFrame을 가공하여 필요한 필드를 추가 및 포맷팅합니다.

    Args:
        df (pd.DataFrame): 원본 데이터프레임
        config (dict): 설정 값

    Returns:
        pd.DataFrame: 가공된 데이터프레임
    """
    df[config["const_sanhak"]] = config["sanhak_name"]
    df["성명(본문)"] = df["성명"].apply(lambda x: Josa.get_full_string(x, "을"))
    df["책임교수명(본문)"] = df["책임교수명"].apply(lambda x: Josa.get_full_string(x, "을"))
    df["총 사업기간"] = df["총 사업기간 시작"] + " ~ " + df["총 사업기간 종료"]
    df["당해연도 사업기간"] = (
        df["당해연도 사업기간 시작"] + " ~ " + df["당해연도 사업기간 종료"]
    )
    df["총 계약금액"] = df["총 계약금액"].astype(int).apply(lambda x: f"{x:,}")
    df["월 계약금액"] = df["월 계약금액"].astype(int).apply(lambda x: f"{x:,}")

    return df
