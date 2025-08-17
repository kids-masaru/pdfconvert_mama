import streamlit as st
import pandas as pd
import io
import os
import re
from openpyxl import load_workbook
import traceback
from pdf_utils import (
    safe_write_df, pdf_to_excel_data_for_paste_sheet, extract_table_from_pdf_for_bento,
    find_correct_anchor_for_bento, extract_bento_range_for_bento, match_bento_names,
    extract_detailed_client_info_from_pdf, export_detailed_client_data_to_dataframe
)

# ページ設定 (アプリ全体に適用)
st.set_page_config(
    page_title="PDF変換ツール",
    page_icon="./static/favicon.ico",
    layout="centered",
)

# --- Session Stateの初期化 ---
def load_master_data(file_path, default_columns):
    if os.path.exists(file_path):
        encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
        for encoding in encodings:
            try:
                df = pd.read_csv(file_path, encoding=encoding)
                if not df.empty:
                    return df
            except Exception:
                continue
    return pd.DataFrame(columns=default_columns)

if 'master_df' not in st.session_state:
    st.session_state.master_df = load_master_data("商品マスタ一覧.csv", ['商品予定名', 'パン箱入数', '商品名'])
if 'customer_master_df' not in st.session_state:
    st.session_state.customer_master_df = load_master_data("得意先マスタ一覧.csv", ['得意先コード', '得意先名'])

# --- サイドバーの見た目を制御 ---
st.markdown("""
    <style>
        /* Streamlitが自動生成するサイドバーの項目を非表示にする */
        [data-testid="stSidebarNav"] ul {
            display: none;
        }
        /* タイトルのデザイン */
        .custom-title {
            font-size: 2.1rem;
            font-weight: 600;
            color: #3A322E;
            padding-bottom: 10px;
