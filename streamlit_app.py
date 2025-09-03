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
    page_icon="./static/icons/android-chrome-192.png",
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
    st.session_state.customer_master_df = load_master_data("得意先マスタ一覧.csv", ['得意先ＣＤ', '得意先名'])

# --- PWAメタタグとサイドバーの見た目を制御 ---
st.markdown("""
    <!-- PWA メタタグ -->
    <link rel="manifest" href="./static/manifest.json">
    <meta name="theme-color" content="#ffffff">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="default">
    <meta name="apple-mobile-web-app-title" content="PDF変換ツール">
    <link rel="apple-touch-icon" href="./static/icons/apple-touch-icon.png">
    <link rel="icon" type="image/png" sizes="192x192" href="./static/icons/android-chrome-192.png">
    <link rel="icon" type="image/png" sizes="512x512" href="./static/icons/android-chrome-512.png">
    
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
            border-bottom: 3px solid #FF9933;
            margin-bottom: 25px;
        }
        .stApp { 
            background: #fff5e6; 
        }
    </style>
""", unsafe_allow_html=True)

st.sidebar.title("メニュー")
# 手動でナビゲーションリンクを作成する
st.sidebar.page_link("streamlit_app.py", label="PDF Excel 変換", icon="📄")
st.sidebar.page_link("pages/マスタ設定.py", label="マスタ設定", icon="⚙️")

# --- ここから下が「PDF→Excel変換」ページのコンテンツ ---
st.markdown('<p class="custom-title">数出表 PDF変換ツール</p>', unsafe_allow_html=True)
uploaded_pdf = st.file_uploader("処理するPDFファイルをアップロードしてください", type="pdf", label_visibility="collapsed")

if uploaded_pdf is not None:
    template_path = "template.xlsm"
    nouhinsyo_path = "nouhinsyo.xlsx"
    if not os.path.exists(template_path) or not os.path.exists(nouhinsyo_path):
        st.error(f"'{template_path}' または '{nouhinsyo_path}' が見つかりません。")
        st.stop()
    
    template_wb = load_workbook(template_path, keep_vba=True)
    nouhinsyo_wb = load_workbook(nouhinsyo_path)
    pdf_bytes_io = io.BytesIO(uploaded_pdf.getvalue())
    
    df_paste_sheet, df_bento_sheet, df_client_sheet = None, None, None
    with st.spinner("PDFからデータを抽出中..."):
        df_paste_sheet = pdf_to_excel_data_for_paste_sheet(io.BytesIO(pdf_bytes_io.getvalue()))
        if df_paste_sheet is not None:
            try:
                tables = extract_table_from_pdf_for_bento(io.BytesIO(pdf_bytes_io.getvalue()))
                if tables:
                    main_table = max(tables, key=len)
                    anchor_col = find_correct_anchor_for_bento(main_table)
                    if anchor_col != -1:
                        bento_list = extract_bento_range_for_bento(main_table, anchor_col)
                        if bento_list:
                            matched_list = match_bento_names(bento_list, st.session_state.master_df)
                            output_data = []
                            for item in matched_list:
                                match = re.search(r' \(入数: (.+?)\)$', item)
                                if match:
                                    output_data.append([item[:match.start()], match.group(1)])
                                else:
                                    output_data.append([item.replace(" (未マッチ)", ""), ""])
                            df_bento_sheet = pd.DataFrame(output_data, columns=['商品予定名', 'パン箱入数'])
            except Exception as e: st.error(f"注文弁当データ処理中にエラー: {e}")
            
            try:
                client_data = extract_detailed_client_info_from_pdf(io.BytesIO(pdf_bytes_io.getvalue()))
                if client_data:
                    df_client_sheet = export_detailed_client_data_to_dataframe(client_data)
                    st.success(f"クライアント情報 {len(client_data)} 件を抽出しました")
            except Exception as e: st.error(f"クライアント情報抽出中にエラー: {e}")
    
    if df_paste_sheet is not None:
        try:
            with st.spinner("Excelファイルを作成中..."):
                ws_paste = template_wb["貼り付け用"]
                for r_idx, row in df_paste_sheet.iterrows():
                    for c_idx, value in enumerate(row):
                        ws_paste.cell(row=r_idx + 1, column=c_idx + 1, value=value)
                if df_bento_sheet is not None: safe_write_df(template_wb["注文弁当の抽出"], df_bento_sheet, start_row=1)
                if df_client_sheet is not None: safe_write_df(template_wb["クライアント抽出"], df_client_sheet, start_row=1)
                output_macro = io.BytesIO()
                template_wb.save(output_macro)
                macro_excel_bytes = output_macro.getvalue()

                df_bento_for_nouhin = None
                if df_bento_sheet is not None:
                    master_df = st.session_state.master_df
                    master_map = master_df.drop_duplicates(subset=['商品予定名']).set_index('商品予定名')['商品名'].to_dict()
                    df_bento_for_nouhin = df_bento_sheet.copy()
                    df_bento_for_nouhin['商品名'] = df_bento_for_nouhin['商品予定名'].map(master_map)
                    df_bento_for_nouhin = df_bento_for_nouhin[['商品予定名', 'パン箱入数', '商品名']]
                
                # nouhinsyo.xlsxへの書き込み処理
                ws_paste_n = nouhinsyo_wb["貼り付け用"]
                for r_idx, row in df_paste_sheet.iterrows():
                    for c_idx, value in enumerate(row):
                        ws_paste_n.cell(row=r_idx + 1, column=c_idx + 1, value=value)
                if df_bento_for_nouhin is not None: 
                    safe_write_df(nouhinsyo_wb["注文弁当の抽出"], df_bento_for_nouhin, start_row=1)
                if df_client_sheet is not None: 
                    safe_write_df(nouhinsyo_wb["クライアント抽出"], df_client_sheet, start_row=1)
                
                # 得意先マスタの書き込みを追加
                if not st.session_state.customer_master_df.empty:
                    safe_write_df(nouhinsyo_wb["得意先マスタ"], st.session_state.customer_master_df, start_row=1)
                
                output_data_only = io.BytesIO()
                nouhinsyo_wb.save(output_data_only)
                data_only_excel_bytes = output_data_only.getvalue()

            st.success("✅ ファイルの準備が完了しました！")
            original_pdf_name = os.path.splitext(uploaded_pdf.name)[0]
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(label="▼　数出表ダウンロード",data=macro_excel_bytes,file_name=f"{original_pdf_name}_数出表.xlsm",mime="application/vnd.ms-excel.sheet.macroEnabled.12")
            with col2:
                st.download_button(label="▼　納品書ダウンロード",data=data_only_excel_bytes,file_name=f"{original_pdf_name}_納品書.xlsx",mime="application/vnd.openxmlformats-officedocument.sheet")
        except Exception as e:
            st.error(f"Excelファイル生成中にエラーが発生しました: {e}")
            traceback.print_exc()
