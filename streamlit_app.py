import streamlit as st
import pandas as pd
import io
import os
import re
from openpyxl import load_workbook
import glob # globモジュールを追加
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
def load_master_data(file_prefix, default_columns):
    # 指定されたプレフィックスで始まるCSVファイルを検索
    # os.path.joinを使ってパスを安全に結合
    list_of_files = glob.glob(os.path.join('.', f'{file_prefix}*.csv'))
    
    # ファイルが見つからなかった場合
    if not list_of_files:
        return pd.DataFrame(columns=default_columns)

    # タイムスタンプ（最終更新日）でソートし、最新のファイルを選択
    latest_file = max(list_of_files, key=os.path.getmtime)
    
    encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
    for encoding in encodings:
        try:
            df = pd.read_csv(latest_file, encoding=encoding)
            if not df.empty:
                return df
        except Exception:
            continue
            
    # 全てのエンコーディングで読み込み失敗した場合
    return pd.DataFrame(columns=default_columns)

if 'master_df' not in st.session_state:
    # ファイルのプレフィックスを指定
    st.session_state.master_df = load_master_data("商品マスタ一覧", ['商品予定名', 'パン箱入数', '商品名'])
if 'customer_master_df' not in st.session_state:
    # ファイルのプレフィックスを指定
    st.session_state.customer_master_df = load_master_data("得意先マスタ一覧", ['得意先ＣＤ', '得意先名'])


# --- PWAメタタグとサイドバーの見た目を制御 ---
st.markdown("""
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
        st.error(f"必要なテンプレートファイルが見つかりません：'{template_path}' または '{nouhinsyo_path}'")
        st.stop()
    
    template_wb = load_workbook(template_path, keep_vba=True)
    nouhinsyo_wb = load_workbook(nouhinsyo_path)
    pdf_bytes_io = io.BytesIO(uploaded_pdf.getvalue())
    
    df_paste_sheet, df_bento_sheet, df_client_sheet = None, None, None
    with st.spinner("PDFからデータを抽出中..."):
        try:
            df_paste_sheet = pdf_to_excel_data_for_paste_sheet(io.BytesIO(pdf_bytes_io.getvalue()))
        except Exception:
            df_paste_sheet = None
            st.error("PDFからの貼り付け用データ抽出中にエラーが発生しました。")

        if df_paste_sheet is not None:
            # 注文弁当データの抽出
            try:
                tables = extract_table_from_pdf_for_bento(io.BytesIO(pdf_bytes_io.getvalue()))
                if tables:
                    main_table = max(tables, key=len)
                    anchor_col = find_correct_anchor_for_bento(main_table)
                    if anchor_col != -1:
                        bento_list = extract_bento_range_for_bento(main_table, anchor_col)
                        if bento_list:
                            # ------------------- ▼ ここから修正 ▼ -------------------
                            matched_list = match_bento_names(bento_list, st.session_state.master_df)
                            output_data = []
                            master_df = st.session_state.master_df

                            # 商品マスタの列数が十分にあるか（R列=18列目まであるか）を確認
                            has_enough_columns = len(master_df.columns) > 17

                            # P列(16列目)とR列(18列目)のヘッダー名を取得。なければデフォルト名を設定
                            col_p_name = master_df.columns[15] if has_enough_columns else '追加データC'
                            col_r_name = master_df.columns[17] if has_enough_columns else '追加データD'

                            for item in matched_list:
                                # 弁当名と入数を抽出
                                bento_name = ""
                                bento_iri = ""
                                match = re.search(r' \(入数: (.+?)\)$', item)
                                if match:
                                    bento_name = item[:match.start()]
                                    bento_iri = match.group(1)
                                else:
                                    bento_name = item.replace(" (未マッチ)", "")

                                val_p = ""
                                val_r = ""
                                
                                # 商品マスタのD列（商品予定名）で一致する行を検索
                                # D列の列名が'商品予定名'であることを前提としています
                                if '商品予定名' in master_df.columns:
                                    matched_row = master_df[master_df['商品予定名'] == bento_name]
                                    # 一致する行があり、かつマスタに十分な列数が存在する場合
                                    if not matched_row.empty and has_enough_columns:
                                        # 最初に見つかった行のP列(16番目)とR列(18番目)の値を取得
                                        val_p = matched_row.iloc[0, 15]
                                        val_r = matched_row.iloc[0, 17]
                                
                                # A, B, C, D列のデータをリストに追加
                                output_data.append([bento_name, bento_iri, val_p, val_r])
                            
                            # 4列構成でDataFrameを作成
                            df_bento_sheet = pd.DataFrame(output_data, columns=['商品予定名', 'パン箱入数', col_p_name, col_r_name])
                            # ------------------- ▲ ここまで修正 ▲ -------------------
            except Exception:
                st.error("注文弁当データ処理中にエラーが発生しました。")

            # クライアント情報の抽出
            try:
                client_data = extract_detailed_client_info_from_pdf(io.BytesIO(pdf_bytes_io.getvalue()))
                if client_data:
                    df_client_sheet = export_detailed_client_data_to_dataframe(client_data)
                    st.success(f"クライアント情報 {len(client_data)} 件を抽出しました")
            except Exception:
                st.error("クライアント情報抽出中にエラーが発生しました。")
    
    if df_paste_sheet is not None:
        try:
            with st.spinner("Excelファイルを作成中..."):
                ws_paste = template_wb["貼り付け用"]
                for r_idx, row in df_paste_sheet.iterrows():
                    for c_idx, value in enumerate(row):
                        ws_paste.cell(row=r_idx + 1, column=c_idx + 1, value=value)
                if df_bento_sheet is not None:
                    safe_write_df(template_wb["注文弁当の抽出"], df_bento_sheet, start_row=1)
                if df_client_sheet is not None:
                    safe_write_df(template_wb["クライアント抽出"], df_client_sheet, start_row=1)
                output_macro = io.BytesIO()
                template_wb.save(output_macro)
                macro_excel_bytes = output_macro.getvalue()

                df_bento_for_nouhin = None
                if df_bento_sheet is not None:
                    master_df = st.session_state.master_df
                    master_map = master_df.drop_duplicates(subset=['商品予定名']).set_index('商品予定名')['商品名'].to_dict()
                    df_bento_for_nouhin = df_bento_sheet.copy()
                    df_bento_for_nouhin['商品名'] = df_bento_for_nouhin['商品予定名'].map(master_map)
                    # 納品書用は従来通り3列に絞り込む
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
                st.download_button(label="▼　数出表ダウンロード", data=macro_excel_bytes, file_name=f"{original_pdf_name}_数出表.xlsm", mime="application/vnd.ms-excel.sheet.macroEnabled.12")
            with col2:
                st.download_button(label="▼　納品書ダウンロード", data=data_only_excel_bytes, file_name=f"{original_pdf_name}_納品書.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception:
            st.error("Excelファイル生成中にエラーが発生しました。詳細はログを確認してください。")
