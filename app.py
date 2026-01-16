from pathlib import Path
from io import BytesIO
import pandas as pd
import streamlit as st


def excel_col_to_idx(col: str) -> int:
    col = col.strip().upper()
    n = 0
    for ch in col:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Invalid Excel column: {col}")
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _is_true(x) -> bool:
    if pd.isna(x):
        return False
    if isinstance(x, bool):
        return x
    s = str(x).strip().lower()
    return s in {"true", "1", "yes", "y"}


def _is_yes(x) -> bool:
    if pd.isna(x):
        return False
    return str(x).strip().lower() == "yes"


def select_by_letters(df: pd.DataFrame, letters: list[str]) -> pd.DataFrame:
    idxs = [excel_col_to_idx(c) for c in letters]
    if max(idxs) >= df.shape[1]:
        raise ValueError(f"ไฟล์มีจำนวนคอลัมน์ไม่ถึง: ต้องการถึงคอลัมน์ {letters[idxs.index(max(idxs))]}")
    return df.iloc[:, idxs]


LEX_FLAG_COL = "AS"
LEX_KEEP_COLS = ["A", "F", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AV"]

SPX_FLAG_COL = "BI"
SPX_KEEP_COLS = ["A", "S", "V", "W", "Y", "BK", "BU", "BV"]


def build_output(lex_file, spx_file) -> bytes:
    if lex_file is not None:
        lex_df = pd.read_excel(lex_file, dtype=object)
        lex_flag_idx = excel_col_to_idx(LEX_FLAG_COL)
        if lex_flag_idx >= lex_df.shape[1]:
            raise ValueError(f"LEX: ไม่พบคอลัมน์ {LEX_FLAG_COL}")
        lex_mask = lex_df.iloc[:, lex_flag_idx].apply(_is_true)
        lex_out = select_by_letters(lex_df.loc[lex_mask].copy(), LEX_KEEP_COLS)
    else:
        lex_out = pd.DataFrame(columns=LEX_KEEP_COLS)

    if spx_file is not None:
        spx_df = pd.read_excel(spx_file, dtype=object)
        spx_flag_idx = excel_col_to_idx(SPX_FLAG_COL)
        if spx_flag_idx >= spx_df.shape[1]:
            raise ValueError(f"SPX: ไม่พบคอลัมน์ {SPX_FLAG_COL}")
        spx_mask = spx_df.iloc[:, spx_flag_idx].apply(_is_yes)
        spx_out = select_by_letters(spx_df.loc[spx_mask].copy(), SPX_KEEP_COLS)
    else:
        spx_out = pd.DataFrame(columns=SPX_KEEP_COLS)

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        lex_out.to_excel(writer, sheet_name="LEX", index=False)
        spx_out.to_excel(writer, sheet_name="SPX", index=False)
    return bio.getvalue()


st.set_page_config(page_title="รายชื่อลูกค้าขอใบกำกับภาษี online")
st.title("รายชื่อลูกค้าขอใบกำกับภาษี online")

lex_upload = st.file_uploader("อัปโหลดไฟล์ LEX (ต้องเป็น .xlsx)", type=["xlsx"], key="lex")
spx_upload = st.file_uploader("อัปโหลดไฟล์ SPX (ต้องเป็น .xlsx)", type=["xlsx"], key="spx")

if st.button("สร้างไฟล์ (Generate)"):
    if lex_upload is None and spx_upload is None:
        st.error("กรุณาอัปโหลดอย่างน้อย 1 ไฟล์ (LEX หรือ SPX)")
    else:
        try:
            out_bytes = build_output(lex_upload, spx_upload)
            st.download_button(
                label="ดาวน์โหลดไฟล์ tax_invoice_request.xlsx",
                data=out_bytes,
                file_name="tax_invoice_request.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.success("สร้างไฟล์เรียบร้อยแล้ว")
        except Exception as e:
            st.error(str(e))
