import streamlit as st
import pandas as pd
import os

# ページ設定 (アプリ全体に適用)
st.set_page_config(
    page_title="PDF変換ツール",
    page_icon="./static/favicon.ico",
    layout="centered",
)

# --- サイドバー ---
st.sidebar.title("メニュー")

# --- 全ページ共通のCSS ---
st.markdown("""
    <style>
        .custom-title {
            font-size: 2.1rem;
            font-weight: 600;
            color: #3A322E;
            padding-bottom: 10px;
            border-bottom: 3px solid #FF9933;
            margin-bottom: 25px;
        }
        .stApp { 
            background: #fff5e6; 
        }
    </style>
""", unsafe_allow_html=True)

# --- 各ページで共通して使う関数 ---
def load_master_data(file_path, default_columns):
    """CSVマスタデータを読み込む関数"""
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

# --- Session Stateの初期化 ---
# アプリ起動時に一度だけ実行され、全ページで共有されます
if 'master_df' not in st.session_state:
    st.session_state.master_df = load_master_data(
        "商品マスタ一覧.csv", 
        ['商品予定名', 'パン箱入数', '商品名']
    )

if 'customer_master_df' not in st.session_state:
    st.session_state.customer_master_df = load_master_data(
        "得意先マスタ一覧.csv", 
        ['得意先コード', '得意先名']
    )

# --- 起動時に表示するページを自動で切り替え ---
if "page" not in st.session_state:
    st.session_state.page = "PDF Excel 変換"

# このファイルはトップページとして表示させない
if st.session_state.page == "PDF Excel 変換":
    st.switch_page("pages/1_PDF_Excel_変換.py")
