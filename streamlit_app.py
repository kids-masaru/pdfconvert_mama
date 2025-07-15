import streamlit as st
import streamlit.components.v1 as components
import os

# ✅ 最も早く呼び出す必要があるため、ここに配置
st.set_page_config(
    page_title="【数出表】PDF → Excelへの変換",
    page_icon="./static/favicon.ico", # faviconのパス (staticフォルダはstreamlit_app.py と同じ階層にある想定)
    layout="centered",
)

# ──────────────────────────────────────────────
# HTML <head> 埋め込み（PWA用 manifest & 各種アイコン）
# ──────────────────────────────────────────────
components.html(
    """
    <link rel="manifest" href="./static/manifest.json">
    <link rel="icon" href="./static/favicon.ico">
    <link rel="apple-touch-icon" sizes="180x180" href="./static/icons/apple-touch-icon.png">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-title" content="YourAppName">
    """,
    height=0,
)

# ──────────────────────────────────────────────
# CSS／UI スタイル定義 (全ページ共通で適用される)
# ──────────────────────────────────────────────
st.markdown("""
    <style>
        /* Google FontsのInter, Robotoをインポート */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Roboto:wght@300;400;500&display=swap');
        .stApp { background: #fff5e6; font-family: 'Inter', sans-serif; }
        .title { font-size: 1.5rem; font-weight: 600; color: #333; margin-bottom: 5px; }
        .subtitle { font-size: 0.9rem; color: #666; margin-bottom: 25px; }
        [data-testid="stFileUploader"] { background: #fff; border-radius: 10px; border: 1px dashed #d0d0d0; padding: 30px 20px; margin: 20px 0; }
        [data-testid="stFileUploader"] label { display: none; }
        [data-testid="stFileUploader"] section { border: none !important; background: transparent !important; }
        .file-card { background: white; border-radius: 8px; padding: 12px 16px; margin: 15px 0; box-shadow: 0 1px 3px rgba(0,0,0,0.08); display: flex; align-items: center; justify-content: space-between; border: 1px solid #eaeaea; }
        .file-info { display: flex; align-items: center; }
        .file-icon { width: 36px; height: 36px; border-radius: 6px; background-color: #f44336; display: flex; align-items: center; justify-content: center; margin-right: 12px; color: white; font-weight: 500; font-size: 14px; }
        .file-details { display: flex; flex-direction: column; }
        .file-name { font-weight: 500; color: #333; font-size: 0.9rem; margin-bottom: 3px; }
        .file-meta { font-size: 0.75rem; color: #888; }
        .loading-spinner { width: 20px; height: 20px; border: 2px solid rgba(0,0,0,0.1); border-radius: 50%; border-top-color: #ff9933; animation: spin 1s linear infinite; }
        .check-icon { color: #ff9933; font-size: 20px; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .progress-bar { height: 4px; background-color: #e0e0e0; border-radius: 2px; width: 100%; margin-top: 10px; overflow: hidden; }
        .progress-value { height: 100%; background-color: #ff9933; border-radius: 2px; width: 60%; transition: width 0.5s ease-in-out; }
        .download-card { background: white; border-radius: 8px; padding: 16px; margin: 20px 0; box-shadow: 0 2px 5px rgba(0,0,0,0.08); display: flex; align-items: center; justify-content: space-between; border: 1px solid #eaeaea; transition: all 0.2s ease; cursor: pointer; text-decoration: none; color: inherit; }
        .download-card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.12); background-color: #fffaf0; transform: translateY(-2px); }
        .download-info { display: flex; align-items: center; }
        .download-icon { width: 40px; height: 40px; border-radius: 8px; background-color: #ff9933; display: flex; align-items: center; justify-content: center; margin-right: 12px; color: white; font-weight: 500; font-size: 16px; }
        .download-details { display: flex; flex-direction: column; }
        .download-name { font-weight: 500; color: #333; font-size: 0.9rem; margin-bottom: 3px; }
        .download-meta { font-size: 0.75rem; color: #888; }
        .download-button-imitation { background-color: #ff9933; color: white; border: none; border-radius: 6px; padding: 8px 16px; font-size: 0.85rem; font-weight: 500; transition: background-color 0.2s; display: flex; align-items: center; }
        .download-card:hover .download-button-imitation { background-color: #e68a00; }
        .download-button-icon { margin-right: 6px; }
        .separator { height: 1px; background-color: #ffe0b3; margin: 25px 0; }
        .st-emotion-cache-16txt4h { /* サイドバーのパディングを調整 */
            padding-top: 2rem;
        }
        .st-emotion-cache-1f87902 { /* サイドバーの要素間のスペースを調整 */
            gap: 1rem;
        }
    </style>
""", unsafe_allow_html=True)

# メインコンテナ開始
st.markdown('<div class="main-container">', unsafe_allow_html=True)

# サイドバーにナビゲーションを追加
st.sidebar.header("ナビゲーション")
st.sidebar.page_link("pages/1_PDF_to_Excel.py", label="PDF → Excel 変換")
st.sidebar.page_link("pages/2_Master_Data_Settings.py", label="設定（マスタデータ更新）")

st.markdown('<div class="title">ようこそ！</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">サイドバーから機能を選択してください。</div>', unsafe_allow_html=True)

st.info("このアプリは、PDFの注文データをExcelに変換し、マスタデータに基づいた情報で「貼り付け用」と「注文弁当の抽出」の2つのシートを自動更新します。")
# 任意の画像パスを設定 (例: static/overview_image.png など)
# もし画像がない場合は、この行をコメントアウトするか、placeholer画像を使用
st.image("https://via.placeholder.com/600x300.png?text=PDF+to+Excel+App+Overview", caption="アプリの概要イメージ") 

st.markdown("""
### アプリの使い方
1.  **PDF → Excel 変換**: PDFファイルからデータを抽出し、`template.xlsm`の「貼り付け用」と「注文弁当の抽出」シートに自動で書き込み、新しいExcelファイルをダウンロードします。
2.  **設定（マスタデータ更新）**: 商品マスタCSVをアップロードすることで、アプリが使用するマスタデータを更新できます。
""")

st.markdown('</div>', unsafe_allow_html=True)
