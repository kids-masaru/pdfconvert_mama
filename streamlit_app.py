import streamlit as st
import pandas as pd
import io
import os
import re
from openpyxl import load_workbook
import glob

# pdf_utils.py から必要な関数をインポート
from pdf_utils import (
    safe_write_df, pdf_to_excel_data_for_paste_sheet, extract_table_from_pdf_for_bento,
    find_correct_anchor_for_bento, extract_bento_range_for_bento, match_bento_names,
    extract_detailed_client_info_from_pdf, export_detailed_client_data_to_dataframe,
    debug_pdf_content
)

# ページ設定 (アプリ全体に適用)
st.set_page_config(
    page_title="PDF変換ツール",
    page_icon="./static/icons/android-chrome-192.png",
    layout="centered",
)

# --- Session Stateの初期化 ---
def load_master_data(file_prefix, default_columns):
    """
    最新の商品マスタCSVを読み込む。
    - 全ての列を文字列として読み込む
    - 空のセルを空文字に変換
    - 「商品予定名」からスペースを完全除去したマッチング専用列を内部的に作成
    """
    list_of_files = glob.glob(os.path.join('.', f'{file_prefix}*.csv'))
    if not list_of_files:
        return pd.DataFrame(columns=default_columns)

    latest_file = max(list_of_files, key=os.path.getmtime)
    
    encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
    for encoding in encodings:
        try:
            df = pd.read_csv(latest_file, encoding=encoding, dtype=str)
            df = df.fillna('')
            if '商品予定名' in df.columns:
                df['商品予定名'] = df['商品予定名'].str.strip()
                # --- ▼修正点▼ ---
                # 半角・全角スペースを全て除去したマッチング専用の列を追加
                df['商品予定名_normalized'] = df['商品予定名'].str.replace(r'\s+', '', regex=True)
                # --- ▲修正点▲ ---
            if not df.empty:
                return df
        except Exception:
            continue
            
    return pd.DataFrame(columns=default_columns)

if 'master_df' not in st.session_state:
    st.session_state.master_df = load_master_data("商品マスタ一覧", ['商品予定名', 'パン箱入数', '商品名'])
if 'customer_master_df' not in st.session_state:
    st.session_state.customer_master_df = load_master_data("得意先マスタ一覧", ['得意先ＣＤ', '得意先名'])


# --- UI設定 ---
st.markdown("""
    <style>
        [data-testid="stSidebarNav"] ul { display: none; }
        .custom-title {
            font-size: 2.1rem; font-weight: 600; color: #3A322E;
            padding-bottom: 10px; border-bottom: 3px solid #FF9933; margin-bottom: 25px;
        }
        .stApp { background: #fff5e6; }
    </style>
""", unsafe_allow_html=True)

st.sidebar.title("メニュー")
st.sidebar.page_link("streamlit_app.py", label="PDF Excel 変換", icon="📄")
st.sidebar.page_link("pages/マスタ設定.py", label="マスタ設定", icon="⚙️")

st.markdown('<p class="custom-title">数出表 PDF変換ツール</p>', unsafe_allow_html=True)

show_debug = st.sidebar.checkbox("デバッグ情報を表示", value=False)
uploaded_pdf = st.file_uploader("処理するPDFファイルをアップロードしてください", type="pdf", label_visibility="collapsed")


# --- メイン処理 ---
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
        except Exception as e:
            df_paste_sheet = None
            st.error(f"PDFからの貼り付け用データ抽出中にエラーが発生しました: {str(e)}")

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
                            master_df = st.session_state.master_df
                            
                            has_enough_columns = len(master_df.columns) > 18
                            col_p_name = master_df.columns[15] if len(master_df.columns) > 15 else '追加データC'
                            col_s_name = master_df.columns[18] if has_enough_columns else '追加データD'

                            if show_debug:
                                st.write("--- 弁当名マッチング状況 ---")

                            for item in matched_list:
                                bento_name, bento_iri = "", ""
                                match = re.search(r' \(入数: (.+?)\)$', item)
                                if match:
                                    bento_name, bento_iri = item[:match.start()], match.group(1)
                                else:
                                    bento_name = item.replace(" (未マッチ)", "")
                                
                                val_p, val_s = "", ""
                                
                                # --- ▼修正点▼ ---
                                # PDFから抽出した弁当名からも全てのスペースを除去して比較する
                                normalized_bento_name = re.sub(r'\s+', '', bento_name)
                                if '商品予定名_normalized' in master_df.columns:
                                    matched_rows = master_df[master_df['商品予定名_normalized'] == normalized_bento_name]
                                # --- ▲修正点▲ ---
                                    
                                    if not matched_rows.empty and has_enough_columns:
                                        first_row = matched_rows.iloc[0]
                                        val_p = str(first_row.iloc[15])
                                        val_s = str(first_row.iloc[18])
                                        
                                        if show_debug:
                                            st.success(f"✅ マッチ成功: '{bento_name}' (as '{normalized_bento_name}') -> P列='{val_p}', S列='{val_s}'")
                                    else:
                                        if show_debug:
                                            reason = "商品マスタに見つかりません。"
                                            st.warning(f"⚠️ マッチ失敗: '{bento_name}' (as '{normalized_bento_name}') - {reason}")
                                
                                output_data.append([bento_name, bento_iri, val_p, val_s])
                            
                            df_bento_sheet = pd.DataFrame(output_data, columns=['商品予定名', 'パン箱入数', col_p_name, col_s_name])
                            
                            if show_debug:
                                st.write("--- 最終的な弁当データ ---")
                                st.dataframe(df_bento_sheet)

            except Exception as e:
                st.error(f"注文弁当データ処理中にエラーが発生しました: {str(e)}")

            try:
                client_data = extract_detailed_client_info_from_pdf(io.BytesIO(pdf_bytes_io.getvalue()))
                if client_data:
                    df_client_sheet = export_detailed_client_data_to_dataframe(client_data)
            except Exception as e:
                st.error(f"クライアント情報抽出中にエラーが発生しました: {str(e)}")
    
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
                    df_bento_for_nouhin = df_bento_for_nouhin[['商品予定名', 'パン箱入数', '商品名']]
                
                ws_paste_n = nouhinsyo_wb["貼り付け用"]
                for r_idx, row in df_paste_sheet.iterrows():
                    for c_idx, value in enumerate(row):
                        ws_paste_n.cell(row=r_idx + 1, column=c_idx + 1, value=value)
                if df_bento_for_nouhin is not None:
                    safe_write_df(nouhinsyo_wb["注文弁当の抽出"], df_bento_for_nouhin, start_row=1)
                if df_client_sheet is not None:
                    safe_write_df(nouhinsyo_wb["クライアント抽出"], df_client_sheet, start_row=1)
                if not st.session_state.customer_master_df.empty:
                    safe_write_df(nouhinsyo_wb["得意先マスタ"], st.session_state.customer_master_df, start_row=1)
                
                output_data_only = io.BytesIO()
                nouhinsyo_wb.save(output_data_only)
                data_only_excel_bytes = output_data_only.getvalue()

            st.success("✅ ファイルの準備が完了しました！")
            original_pdf_name = os.path.splitext(uploaded_pdf.name)[0]
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="▼　数出表ダウンロード", data=macro_excel_bytes,
                    file_name=f"{original_pdf_name}_数出表.xlsm",
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12"
                )
            with col2:
                st.download_button(
                    label="▼　納品書ダウンロード", data=data_only_excel_bytes,
                    file_name=f"{original_pdf_name}_納品書.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Excelファイル生成中にエラーが発生しました: {str(e)}")
