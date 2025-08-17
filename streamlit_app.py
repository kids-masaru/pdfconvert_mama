import streamlit as st

# ページ設定 (アプリ全体に適用)
st.set_page_config(
    page_title="PDF変換ツール",
    page_icon="./static/favicon.ico",
    layout="centered",
)

# --- サイドバー ---
# これが最初に実行されるため、メニュータイトルが一番上に表示される
st.sidebar.title("メニュー")

# --- 全ページ共通のCSSとコンポーネント ---
# 新しいタイトルのデザインをここで定義
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

# Main app logic can be minimal or a welcome message
st.markdown("サイドバーのメニューから操作を選択してください。")
