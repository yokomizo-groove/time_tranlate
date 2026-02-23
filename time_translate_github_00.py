import streamlit as st
import pandas as pd
import numpy as np
import os
from io import BytesIO

st.set_page_config(page_title="勤怠変換アプリ", layout="centered")
st.title("勤怠データ変換アプリ")


# ===== 時刻変換 =====
def convert_time_series(series):
    s = series.fillna("").astype(str).str.strip()
    hm = s.str.extract(r'^(\d{1,2})[:\'](\d{1,2})$')

    h = pd.to_numeric(hm[0], errors='coerce')
    m = pd.to_numeric(hm[1], errors='coerce')

    result = (h * 100 + m).fillna(0).astype("int32")
    return result


# ===== メイン変換処理 =====
def process_file(uploaded_file):

    ext = os.path.splitext(uploaded_file.name)[1].lower()

    if ext == ".csv":
        try:
            df = pd.read_csv(uploaded_file, dtype=str, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(uploaded_file, dtype=str, encoding="cp932")
    elif ext in [".xlsx", ".xlsm"]:
        df = pd.read_excel(uploaded_file, dtype=str)
    else:
        st.error("対応していないファイル形式です")
        return None

    row_count = len(df)
    MAX_COL = 150

    # numpyへ変換
    base_array = df.to_numpy(dtype=object)

    if base_array.shape[1] < MAX_COL:
        pad = np.empty((row_count, MAX_COL - base_array.shape[1]), dtype=object)
        pad[:] = ""
        base_array = np.hstack([base_array, pad])

    final_array = base_array

    mapping = {
        99: "法定内超勤時間",
        100: "早出残業時間",
        101: "普通残業時間",
        102: "実労働時間",
        103: "所定内深夜時間",
        104: "所定外深夜時間",
        106: "所定外勤務時間",
        107: "休日深夜時間",
        108: "乖離時間（始業）",
        109: "乖離時間（終業）",
        110: "年休換算時間",
        111: "調休換算時間",
        112: "不就業１時間",
        113: "所定内労働時間",
        114: "休憩時間",
        115: "特休勤務時間",
        116: "公休勤務時間",
        121: "出勤打刻",
        122: "退勤打刻",
        123: "始業時刻",
        124: "終業時刻",
    }

    converted_cache = {}

    for excel_col, col_name in mapping.items():
        if col_name in df.columns:
            converted = convert_time_series(df[col_name].iloc[1:])
            final_array[1:, excel_col - 1] = converted.values
            converted_cache[col_name] = converted

    # 深夜時間計
    if "所定内深夜時間" in converted_cache and "所定外深夜時間" in converted_cache:
        total = (
            converted_cache["所定内深夜時間"]
            + converted_cache["所定外深夜時間"]
        )
        final_array[1:, 105 - 1] = total.values

    final_df = pd.DataFrame(final_array)

    # ヘッダー整形
    headers = list(df.columns)
    while len(headers) < final_df.shape[1]:
        headers.append("")

    for excel_col, col_name in mapping.items():
        if excel_col - 1 < len(headers):
            headers[excel_col - 1] = col_name + "-t"

    if 105 - 1 < len(headers):
        headers[105 - 1] = "深夜時間計-t"

    final_df.columns = headers

    # BytesIO出力
    output = BytesIO()
    final_df.to_excel(output, index=False, engine="xlsxwriter")
    output.seek(0)

    return output


# ===== UI =====
uploaded_file = st.file_uploader(
    "CSVまたはExcelファイルをアップロードしてください",
    type=["csv", "xlsx", "xlsm"]
)

if uploaded_file is not None:

    st.success("ファイルを読み込みました")

    result_file = process_file(uploaded_file)

    if result_file:

        base_name = os.path.splitext(uploaded_file.name)[0]
        download_name = f"{base_name}_output.xlsx"

        st.download_button(
            label="変換ファイルをダウンロード",
            data=result_file,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )