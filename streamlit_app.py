import streamlit as st
import streamlit.components.v1 as components
import pdfplumber
import pandas as pd
import io
import re
import os
import unicodedata
import traceback
from typing import List, Dict, Any
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ページ設定
st.set_page_config(
    page_title="【数出表】PDF → Excelへの変換",
    page_icon="./static/favicon.ico",
    layout="centered",
)

# --- Streamlit Session Stateの初期化 ---
if 'master_df' not in st.session_state:
    master_csv_path = "商品マスタ一覧.csv"
    initial_master_df = None
    if os.path.exists(master_csv_path):
        encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis', 'euc-jp', 'iso-2022-jp']
        for encoding in encodings:
            try:
                temp_df = pd.read_csv(master_csv_path, encoding=encoding)
                if not temp_df.empty:
                    initial_master_df = temp_df
                    st.success(f"既存のマスタデータを {encoding} エンコーディングで読み込みました。")
                    break
            except (UnicodeDecodeError, pd.errors.EmptyDataError):
                continue
            except Exception as e:
                st.warning(f"既存マスタCSV ({master_csv_path}) を {encoding} で読み込み中にエラーが発生しました: {e}")
                continue
    if initial_master_df is None:
        st.warning(f"マスタデータ '{master_csv_path}' が見つからないか、読み込めませんでした。マスタ設定ページでアップロードしてください。")
        initial_master_df = pd.DataFrame(columns=['商品予定名', 'パン箱入数', '商品名'])
    st.session_state.master_df = initial_master_df

# --- HTML/CSS, サイドバー ---
components.html("""<link rel="manifest" href="./static/manifest.json">""", height=0)
st.markdown("""<style>.stApp { background: #fff5e6; }</style>""", unsafe_allow_html=True)
st.sidebar.title("メニュー")
page_selection = st.sidebar.radio("表示する機能を選択してください", ("PDF → Excel 変換", "マスタ設定"), index=0)
st.markdown("---")

# ──────────────────────────────────────────────
# ✅【変更点】数式を保護しながら書き込む新しいヘルパー関数
# ──────────────────────────────────────────────
def safe_write_df(worksheet, df, start_row=2):
    """
    数式を保護するため、指定された範囲のセルのみをクリアし、データフレームを書き込む
    """
    num_cols = df.shape[1]
    
    # 1. 既存データのクリア（指定列のみ）
    # 書き込む行数より既存の行数が多い場合、余分な行のデータをクリアする
    if worksheet.max_row >= start_row:
        for row in range(start_row, worksheet.max_row + 1):
            for col in range(1, num_cols + 1):
                worksheet.cell(row=row, column=col).value = None

    # 2. 新しいデータの書き込み（ヘッダーは除く）
    for r_idx, row_data in enumerate(df.itertuples(index=False), start=start_row):
        for c_idx, value in enumerate(row_data, start=1):
            worksheet.cell(row=r_idx, column=c_idx, value=value)

# (PDF解析・データ抽出関数群は変更がないため省略します)
# ... (All the data extraction functions like extract_detailed_client_info_from_pdf, etc. go here) ...
def extract_detailed_client_info_from_pdf(pdf_file_obj):
    """PDFから詳細なクライアント情報（名前＋給食の数）を抽出する"""
    client_data = []

    try:
        with pdfplumber.open(pdf_file_obj) as pdf:
            for page_num, page in enumerate(pdf.pages):
                rows = extract_text_with_layout(page)
                if not rows:
                    continue
                garden_row_idx = -1
                for i, row in enumerate(rows):
                    row_text = ''.join(str(cell) for cell in row if cell)
                    if '園名' in row_text:
                        garden_row_idx = i
                        break
                if garden_row_idx == -1:
                    continue
                current_client_id = None
                current_client_name = None
                for i in range(garden_row_idx + 1, len(rows)):
                    row = rows[i]
                    row_text = ''.join(str(cell) for cell in row if cell)
                    if '10001' in row_text:
                        break
                    if not any(str(cell).strip() for cell in row):
                        continue
                    if len(row) > 0 and row[0]:
                        left_cell = str(row[0]).strip()
                        if re.match(r'^\d+$', left_cell):
                            if current_client_id and current_client_name:
                                client_info = extract_meal_numbers_from_row(rows, i-1, current_client_id, current_client_name)
                                if client_info:
                                    client_data.append(client_info)
                            current_client_id = left_cell
                            current_client_name = None
                        elif not re.match(r'^\d+$', left_cell) and current_client_id:
                            current_client_name = left_cell
                if current_client_id and current_client_name:
                    client_info = extract_meal_numbers_from_row(rows, len(rows)-1, current_client_id, current_client_name)
                    if client_info:
                        client_data.append(client_info)
    except Exception as e:
        st.error(f"クライアント情報抽出中にエラーが発生しました: {e}")
    return client_data

def extract_meal_numbers_from_row(rows, row_idx, client_id, client_name):
    client_info = {'client_id': client_id, 'client_name': client_name, 'student_meals': [], 'teacher_meals': []}
    rows_to_check = []
    id_row_idx = -1
    name_row_idx = -1
    for i in range(max(0, row_idx - 3), min(len(rows), row_idx + 3)):
        if i < len(rows) and len(rows[i]) > 0:
            left_cell = str(rows[i][0]).strip()
            if left_cell == client_id:
                id_row_idx = i
                rows_to_check.append(('id', i, rows[i]))
            elif left_cell == client_name:
                name_row_idx = i
                rows_to_check.append(('name', i, rows[i]))
    all_numbers = []
    for row_type, idx, row in rows_to_check:
        for col_idx in range(1, len(row)):
            cell = str(row[col_idx]).strip()
            if cell and re.match(r'^\d+$', cell):
                all_numbers.append({'number': int(cell), 'row_type': row_type, 'col_idx': col_idx})
            elif cell and not re.match(r'^\d+$', cell) and cell != '':
                break
    id_numbers = [item['number'] for item in all_numbers if item['row_type'] == 'id']
    name_numbers = [item['number'] for item in all_numbers if item['row_type'] == 'name']
    client_info['student_meals'] = id_numbers[:3]
    client_info['teacher_meals'] = name_numbers[:2]
    return client_info

def export_detailed_client_data_to_dataframe(client_data):
    df_data = []
    for client_info in client_data:
        row = {
            'クライアント名': client_info['client_name'],
            '園児の給食の数1': client_info['student_meals'][0] if len(client_info['student_meals']) > 0 else '',
            '園児の給食の数2': client_info['student_meals'][1] if len(client_info['student_meals']) > 1 else '',
            '園児の給食の数3': client_info['student_meals'][2] if len(client_info['student_meals']) > 2 else '',
            '先生の給食の数1': client_info['teacher_meals'][0] if len(client_info['teacher_meals']) > 0 else '',
            '先生の給食の数2': client_info['teacher_meals'][1] if len(client_info['teacher_meals']) > 1 else '',
        }
        df_data.append(row)
    return pd.DataFrame(df_data)

def pdf_to_excel_data_for_paste_sheet(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            if not pdf.pages:
                st.warning("PDFにページがありません。")
                return None
            page = pdf.pages[0]
            rows = extract_text_with_layout(page)
            rows = [row for row in rows if any(cell.strip() for cell in row)]
            if not rows:
                st.warning("PDFの最初のページからテキストデータを抽出できませんでした。")
                return None
            rows = post_process_rows(rows)
            rows = remove_extra_empty_columns(rows)
            if not rows or not rows[0]:
                st.warning("空の列を削除した結果、データがなくなりました。")
                return None
            max_cols = max(len(row) for row in rows) if rows else 0
            normalized_rows = [row + [''] * (max_cols - len(row)) for row in rows]
            df = pd.DataFrame(normalized_rows)
            return df
    except Exception as e:
        st.error(f"PDF処理中にエラーが発生しました: {e}")
        return None

def extract_table_from_pdf_for_bento(pdf_file_obj):
    tables = []
    with pdfplumber.open(pdf_file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text is None: continue
            if not any(kw in text for kw in ["園名", "飯あり", "キャラ弁"]):
                continue
            lines = page.lines
            if not lines: continue
            y_coords = sorted(set([line['top'] for line in lines] + [line['bottom'] for line in lines]))
            if len(y_coords) < 2: continue
            table_top = min(y_coords)
            table_bottom = max(y_coords)
            x_coords = sorted(set([line['x0'] for line in lines] + [line['x1'] for line in lines]))
            if len(x_coords) < 2: continue
            table_left = min(x_coords)
            table_right = max(x_coords)
            table_bbox = (table_left, table_top, table_right, table_bottom)
            cropped_page = page.crop(table_bbox)
            table_settings = {"vertical_strategy": "lines", "horizontal_strategy": "lines", "snap_tolerance": 3, "join_tolerance": 3, "edge_min_length": 15}
            table = cropped_page.extract_table(table_settings)
            if table:
                tables.append(table)
    return tables

def find_correct_anchor_for_bento(table, target_row_text="赤"):
    for row_idx, row in enumerate(table):
        row_text = ''.join(str(cell) for cell in row if cell)
        if target_row_text in row_text:
            for offset in [1, 2]:
                if row_idx + offset < len(table):
                    next_row = table[row_idx + offset]
                    for col_idx, cell in enumerate(next_row):
                        if cell and "飯なし" in cell:
                            return col_idx
    return -1

def extract_bento_range_for_bento(table, start_col):
    bento_list = []
    end_col = -1
    for row in table:
        row_text = ''.join(str(cell) for cell in row if cell)
        if "おやつ" in row_text:
            for col_idx, cell in enumerate(row):
                if cell and "おやつ" in cell:
                    end_col = col_idx
                    break
            if end_col != -1:
                break
    if end_col == -1 or start_col >= end_col:
        return []
    header_row_idx = None
    anchor_row_idx = -1
    for row_idx, row in enumerate(table):
        if any(cell and "飯なし" in cell for cell in row):
            anchor_row_idx = row_idx
            break
    if anchor_row_idx == -1: return []
    if anchor_row_idx - 1 >= 0:
        header_row_idx = anchor_row_idx - 1
    else:
        return []
    for col in range(start_col + 1, end_col + 1):
        cell_text = table[header_row_idx][col] if col < len(table[header_row_idx]) else ""
        if cell_text and str(cell_text).strip() and "飯なし" not in str(cell_text):
            bento_list.append(str(cell_text).strip())
    return bento_list

def match_bento_names(pdf_bento_list, master_df):
    if master_df is None or master_df.empty:
        st.error("マスタデータがロードされていません。")
        return [f"{name} (マスタデータなし)" for name in pdf_bento_list]
    master_data_tuples = []
    try:
        if '商品予定名' in master_df.columns and 'パン箱入数' in master_df.columns:
            master_data_tuples = master_df[['商品予定名', 'パン箱入数']].dropna().values.tolist()
            master_data_tuples = [(str(name), str(value)) for name, value in master_data_tuples]
        else:
            return [f"{name} (必要列なし)" for name in pdf_bento_list]
    except Exception as e:
        st.error(f"マスタデータ処理中にエラーが発生しました: {e}")
        return [f"{name} (処理エラー)" for name in pdf_bento_list]

    matched = []
    normalized_master_data_tuples = []
    for master_name, master_id in master_data_tuples:
        normalized_name = unicodedata.normalize('NFKC', master_name)
        normalized_name = re.sub(r'\s+', '', normalized_name)
        normalized_master_data_tuples.append((normalized_name, master_name, master_id))

    for pdf_name in pdf_bento_list:
        original_normalized_pdf_name = unicodedata.normalize('NFKC', str(pdf_name))
        original_normalized_pdf_name = re.sub(r'\s+', '', original_normalized_pdf_name)
        found_match = False
        found_original_master_name = None
        found_id = None
        for norm_m_name, orig_m_name, m_id in normalized_master_data_tuples:
            if norm_m_name.startswith(original_normalized_pdf_name):
                found_original_master_name = orig_m_name
                found_id = m_id
                found_match = True
                break
        if not found_match:
            for norm_m_name, orig_m_name, m_id in normalized_master_data_tuples:
                if original_normalized_pdf_name in norm_m_name:
                    found_original_master_name = orig_m_name
                    found_id = m_id
                    found_match = True
                    break
        if found_original_master_name:
            if found_id:
                matched.append(f"{found_original_master_name} (入数: {found_id})")
            else:
                matched.append(found_original_master_name)
        else:
            matched.append(f"{pdf_name} (未マッチ)")
    return matched
def extract_text_with_layout(page) -> List[List[str]]:
    words = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False)
    if not words: return []
    boundaries = get_vertical_boundaries(page)
    if len(boundaries) < 2:
        lines = page.extract_text(layout=False, x_tolerance=3, y_tolerance=3)
        return [[line] for line in lines.split('\n') if line.strip()]
    row_groups = get_line_groups(words, y_tolerance=1.5)
    result_rows = []
    for group in row_groups:
        sorted_group = sorted(group, key=lambda w: w['x0'])
        columns = split_line_using_boundaries(sorted_group, boundaries)
        if any(cell.strip() for cell in columns):
            result_rows.append(columns)
    return result_rows
def remove_extra_empty_columns(rows: List[List[str]]) -> List[List[str]]:
    if not rows: return rows
    num_cols = max(len(row) for row in rows) if rows else 0
    if num_cols == 0: return rows
    is_col_empty = [True] * num_cols
    for r, row in enumerate(rows):
        for c in range(len(row)):
            if c < num_cols and row[c].strip():
                is_col_empty[c] = False
    keep_indices = [c for c in range(num_cols) if not is_col_empty[c]]
    new_rows = []
    for row in rows:
        new_row = [row[i] if i < len(row) else "" for i in keep_indices]
        new_rows.append(new_row)
    return new_rows
def post_process_rows(rows: List[List[str]]) -> List[List[str]]:
    new_rows = [row[:] for row in rows]
    for i, row in enumerate(new_rows):
        for j, cell in enumerate(row):
            if "合計" in str(cell):
                if i > 0 and j < len(new_rows[i-1]):
                    new_rows[i-1][j] = ""
    return new_rows


# ──────────────────────────────────────────────
# メインアプリケーション
# ──────────────────────────────────────────────
if page_selection == "PDF → Excel 変換":
    st.markdown('<div class="title">【数出表】PDF → Excelへの変換</div>', unsafe_allow_html=True)

    uploaded_pdf = st.file_uploader("処理するPDFファイルをアップロードしてください", type="pdf")

    if uploaded_pdf is not None:
        template_path = "template.xlsm"
        nouhinsyo_path = "nouhinsyo.xlsx"
        if not os.path.exists(template_path) or not os.path.exists(nouhinsyo_path):
            st.error(f"'{template_path}' または '{nouhinsyo_path}' が見つかりません。")
            st.stop()
        
        # --- 処理のたびにテンプレートをメモリに読み込む ---
        template_wb = load_workbook(template_path, keep_vba=True)
        nouhinsyo_wb = load_workbook(nouhinsyo_path)

        pdf_bytes_io = io.BytesIO(uploaded_pdf.getvalue())
        
        # --- データ抽出 ---
        df_paste_sheet, df_bento_sheet, df_client_sheet = None, None, None
        with st.spinner("PDFからデータを抽出中..."):
            df_paste_sheet = pdf_to_excel_data_for_paste_sheet(io.BytesIO(pdf_bytes_io.getvalue()))
            if df_paste_sheet is not None:
                # 弁当抽出
                try:
                    tables = extract_table_from_pdf_for_bento(io.BytesIO(pdf_bytes_io.getvalue()))
                    if tables:
                        main_table = max(tables, key=lambda t: len(t) * len(t[0]))
                        if main_table:
                            anchor_col = find_correct_anchor_for_bento(main_table)
                            if anchor_col != -1:
                                bento_list = extract_bento_range_for_bento(main_table, anchor_col)
                                if bento_list:
                                    matched_list = match_bento_names(bento_list, st.session_state.master_df)
                                    output_data_bento = []
                                    for item in matched_list:
                                        match = re.search(r' \(入数: (.+?)\)$', item)
                                        if match:
                                            bento_name, bento_count = item[:match.start()], match.group(1)
                                            output_data_bento.append([bento_name.strip(), bento_count.strip()])
                                        elif "(未マッチ)" in item:
                                            output_data_bento.append([item.replace(" (未マッチ)", "").strip(), ""])
                                        else:
                                            output_data_bento.append([item.strip(), ""])
                                    df_bento_sheet = pd.DataFrame(output_data_bento, columns=['商品予定名', 'パン箱入数'])
                except Exception as e:
                    st.error(f"注文弁当データ処理中にエラー: {e}")
                
                # クライアント抽出
                try:
                    client_data = extract_detailed_client_info_from_pdf(io.BytesIO(pdf_bytes_io.getvalue()))
                    if client_data:
                        df_client_sheet = export_detailed_client_data_to_dataframe(client_data)
                        st.success(f"クライアント情報 {len(client_data)} 件を抽出しました")
                except Exception as e:
                    st.error(f"クライアント情報抽出中にエラー: {e}")
        
        if df_paste_sheet is not None:
            try:
                # --- template.xlsmへの書き込み ---
                with st.spinner("template.xlsm を作成中..."):
                    # 貼り付け用シート (ここはセル指定なので従来通り)
                    ws_paste = template_wb["貼り付け用"]
                    for r_idx, row in df_paste_sheet.iterrows():
                        for c_idx, value in enumerate(row):
                            ws_paste.cell(row=r_idx + 1, column=c_idx + 1, value=value)
                    
                    # 注文弁当の抽出 (A,B列のみ書き込み)
                    if df_bento_sheet is not None:
                        ws_bento = template_wb["注文弁当の抽出"]
                        safe_write_df(ws_bento, df_bento_sheet, start_row=2)
                    
                    # クライアント抽出 (A-F列のみ書き込み)
                    if df_client_sheet is not None:
                        ws_client = template_wb["クライアント抽出"]
                        safe_write_df(ws_client, df_client_sheet, start_row=2)

                    output_macro = io.BytesIO()
                    template_wb.save(output_macro)
                    macro_excel_bytes = output_macro.getvalue()

                # --- nouhinsyo.xlsxへの書き込み ---
                with st.spinner("nouhinsyo.xlsx を作成中..."):
                    df_bento_for_nouhin = None
                    if df_bento_sheet is not None:
                        master_df = st.session_state.master_df
                        master_map = master_df.drop_duplicates(subset=['商品予定名']).set_index('商品予定名')['商品名'].to_dict()
                        df_bento_for_nouhin = df_bento_sheet.copy()
                        df_bento_for_nouhin['商品名'] = df_bento_for_nouhin['商品予定名'].map(master_map)
                        df_bento_for_nouhin = df_bento_for_nouhin[['商品予定名', 'パン箱入数', '商品名']]

                    # 貼り付け用シート
                    ws_paste_n = nouhinsyo_wb["貼り付け用"]
                    for r_idx, row in df_paste_sheet.iterrows():
                        for c_idx, value in enumerate(row):
                            ws_paste_n.cell(row=r_idx + 1, column=c_idx + 1, value=value)
                    
                    # 注文弁当の抽出 (A,B,C列のみ書き込み)
                    if df_bento_for_nouhin is not None:
                        ws_bento_n = nouhinsyo_wb["注文弁当の抽出"]
                        safe_write_df(ws_bento_n, df_bento_for_nouhin, start_row=2)
                    
                    # クライアント抽出
                    if df_client_sheet is not None:
                        ws_client_n = nouhinsyo_wb["クライアント抽出"]
                        safe_write_df(ws_client_n, df_client_sheet, start_row=2)
                    
                    output_data_only = io.BytesIO()
                    nouhinsyo_wb.save(output_data_only)
                    data_only_excel_bytes = output_data_only.getvalue()

                # --- ダウンロードボタンの表示 ---
                st.success("✅ ファイルの準備が完了しました！")
                original_pdf_name = os.path.splitext(uploaded_pdf.name)[0]
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📥 マクロ付きExcelをダウンロード",
                        data=macro_excel_bytes,
                        file_name=f"{original_pdf_name}_Processed.xlsm",
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    )
                with col2:
                    st.download_button(
                        label="📥 データExcelをダウンロード",
                        data=data_only_excel_bytes,
                        file_name=f"{original_pdf_name}_Data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except Exception as e:
                st.error(f"Excelファイル生成中にエラーが発生しました: {e}")
                traceback.print_exc()

# マスタ設定 ページ (変更なし)
elif page_selection == "マスタ設定":
    # (省略)
    st.markdown('<div class="title">マスタデータ設定</div>', unsafe_allow_html=True)
    master_csv_path = "商品マスタ一覧.csv"
    st.markdown("#### 新しいマスタをアップロード")
    uploaded_master_csv = st.file_uploader(
        "商品マスタ一覧.csv をアップロードしてください",
        type="csv",
        help="ヘッダーには '商品予定名', 'パン箱入数', '商品名' を含めてください。"
    )
    if uploaded_master_csv is not None:
        try:
            new_master_df = None
            encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
            for encoding in encodings:
                try:
                    uploaded_master_csv.seek(0)
                    temp_df = pd.read_csv(uploaded_master_csv, encoding=encoding)
                    if all(col in temp_df.columns for col in ['商品予定名', 'パン箱入数', '商品名']):
                        new_master_df = temp_df
                        st.info(f"ファイルを {encoding} で読み込みました。")
                        break
                    else:
                        st.warning(f"{encoding} で読み込みましたが、必須列が見つかりません。")
                except (UnicodeDecodeError, pd.errors.ParserError):
                    continue
                except Exception as e:
                    st.error(f"読み込み中にエラー: {e}")
                    break
            if new_master_df is not None:
                st.session_state.master_df = new_master_df
                try:
                    new_master_df.to_csv(master_csv_path, index=False, encoding='utf-8-sig')
                    st.success(f"✅ マスタデータを更新し、'{master_csv_path}' に保存しました。")
                except Exception as e:
                    st.error(f"マスタファイル保存中にエラー: {e}")
            else:
                st.error("CSVファイルを正しく読み込めませんでした。")
        except Exception as e:
            st.error(f"マスタ更新処理中にエラー: {e}")
    st.markdown("#### 現在のマスタデータ")
    if 'master_df' in st.session_state and not st.session_state.master_df.empty:
        st.dataframe(st.session_state.master_df, use_container_width=True)
    else:
        st.warning("現在、マスタデータが読み込まれていません。")
