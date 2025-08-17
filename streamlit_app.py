import streamlit as st
import pandas as pd
import os

# ページ設定 (アプリ全体に適用)
st.set_page_config(
    page_title="【数出表】PDF → Excelへの変換",
    page_icon="./static/favicon.ico",
    layout="centered",
)

# --- Streamlit Session Stateの初期化 ---
# この処理は、どのページに移動しても最初に実行されるため、ここに置くのが最適です。

# 商品マスタの読み込み
if 'master_df' not in st.session_state:
    master_csv_path = "商品マスタ一覧.csv"
    initial_master_df = None
    if os.path.exists(master_csv_path):
        encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
        for encoding in encodings:
            try:
                temp_df = pd.read_csv(master_csv_path, encoding=encoding)
                if not temp_df.empty:
                    initial_master_df = temp_df
                    break
            except Exception:
                continue
    if initial_master_df is None:
        initial_master_df = pd.DataFrame(columns=['商品予定名', 'パン箱入数', '商品名'])
    st.session_state.master_df = initial_master_df

# 得意先マスタの読み込み
if 'customer_master_df' not in st.session_state:
    customer_master_csv_path = "得意先マスタ一覧.csv"
    initial_customer_master_df = None
    if os.path.exists(customer_master_csv_path):
        encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
        for encoding in encodings:
            try:
                temp_df = pd.read_csv(customer_master_csv_path, encoding=encoding)
                if not temp_df.empty:
                    initial_customer_master_df = temp_df
                    break
            except Exception:
                continue
    if initial_customer_master_df is None:
        initial_customer_master_df = pd.DataFrame(columns=['得意先コード', '得意先名'])
    st.session_state.customer_master_df = initial_customer_master_df

# --- アプリのメインタイトル ---
st.sidebar.title("メニュー")
st.markdown("### 数出表 PDF変換ツール")
st.markdown("サイドバーのメニューから操作を選択してください。")

# --- 全ページ共通のCSSとコンポーネント ---
st.markdown("""<style>.stApp { background: #fff5e6; }</style>""", unsafe_allow_html=True)
