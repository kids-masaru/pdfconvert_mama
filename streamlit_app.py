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
# 商品マスタの読み込み
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
                    st.success(f"既存の商品マスタを {encoding} で読み込みました。")
                    break
            except Exception:
                continue
    if initial_master_df is None:
        st.warning(f"商品マスタ '{master_csv_path}' が見つからないか、読み込めませんでした。")
        initial_master_df = pd.DataFrame(columns=['商品予定名', 'パン箱入数', '商品名'])
    st.session_state.master_df = initial_master_df

# 得意先マスタの読み込み
if 'customer_master_df' not in st.session_state:
    customer_master_csv_path = "得意先マスタ一覧.csv" # ✅ ファイル名を修正
    initial_customer_master_df = None
    if os.path.exists(customer_master_csv_path):
        encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
        for encoding in encodings:
            try:
                temp_df = pd.read_csv(customer_master_csv_path, encoding=encoding)
                if not temp_df.empty:
                    initial_customer_master_df = temp_df
                    st.success(f"既存の得意先マスタを {encoding} で読み込みました。")
                    break
            except Exception:
                continue
    if initial_customer_master_df is None:
        st.warning(f"得意先マスタ '{customer_master_csv_path}' が見つからないか、読み込めませんでした。")
        initial_customer_master_df = pd.DataFrame(columns=['得意先コード', '得意先名'])
    st.session_state.customer_master_df = initial_customer_master_df


# --- HTML/CSS, サイドバー ---
components.html("""<link rel="manifest" href="./static/manifest.json">""", height=0)
st.markdown("""<style>.stApp { background: #fff5e6; }</style>""", unsafe_allow_html=True)
st.sidebar.title("メニュー")
page_selection = st.sidebar.radio("表示する機能を選択してください", ("PDF → Excel 変換", "マスタ設定"), index=0)
st.markdown("---")

# ──────────────────────────────────────────────
# ヘルパー関数
# ──────────────────────────────────────────────
def safe_write_df(worksheet, df, start_row=1):
    num_cols = df.shape[1]
    if worksheet.max_row >= start_row:
        for row in range(start_row, worksheet.max_row + 1):
            for col in range(1, num_cols + 1):
                worksheet.cell(row=row, column=col).value = None
    for r_idx, row_data in enumerate(df.itertuples(index=False), start=start_row):
        for c_idx, value in enumerate(row_data, start=1):
            worksheet.cell(row=r_idx, column=c_idx, value=value)

# ──────────────────────────────────────────────
# PDF解析・データ抽出関数群 (内容は変更なし)
# ──────────────────────────────────────────────
def extract_detailed_client_info_from_pdf(pdf_file_obj):
    client_data = []
    try:
        with pdfplumber.open(pdf_file_obj) as pdf:
            for page in pdf.pages:
                rows = extract_text_with_layout(page)
                if not rows: continue
                garden_row_idx = -1
                for i, row in enumerate(rows):
                    if '園名' in ''.join(str(c) for c in row if c):
                        garden_row_idx = i
                        break
                if garden_row_idx == -1: continue
                current_client_id, current_client_name = None, None
                for i in range(garden_row_idx + 1, len(rows)):
                    row = rows[i]
                    row_text = ''.join(str(c) for c in row if c)
                    if '10001' in row_text: break
                    if not any(str(c).strip() for c in row): continue
                    if row and row[0]:
                        left_cell = str(row[0]).strip()
                        if re.match(r'^\d+$', left_cell):
                            if current_client_id and current_client_name:
                                client_info = extract_meal_numbers_from_row(rows, i - 1, current_client_id, current_client_name)
                                if client_info: client_data.append(client_info)
                            current_client_id, current_client_name = left_cell, None
                        elif not re.match(r'^\d+$', left_cell) and current_client_id:
                            current_client_name = left_cell
                if current_client_id and current_client_name:
                    client_info = extract_meal_numbers_from_row(rows, len(rows) - 1, current_client_id, current_client_name)
                    if client_info: client_data.append(client_info)
    except Exception as e:
        st.error(f"クライアント情報抽出中にエラー: {e}")
    return client_data

def extract_meal_numbers_from_row(rows, row_idx, client_id, client_name):
    client_info = {'client_id': client_id, 'client_name': client_name, 'student_meals': [], 'teacher_meals': []}
    rows_to_check = []
    for i in range(max(0, row_idx - 3), min(len(rows), row_idx + 3)):
        if i < len(rows) and rows[i]:
            left_cell = str(rows[i][0]).strip()
            if left_cell == client_id: rows_to_check.append(('id', i, rows[i]))
            elif left_cell == client_name: rows_to_check.append(('name', i, rows[i]))
    all_numbers = []
    for row_type, _, row in rows_to_check:
        for col_idx, cell in enumerate(row[1:], 1):
            cell_str = str(cell).strip()
            if cell_str and re.match(r'^\d+$', cell_str):
                all_numbers.append({'number': int(cell_str), 'row_type': row_type})
            elif cell_str and not re.match(r'^\d+$', cell_str):
                break
    client_info['student_meals'] = [item['number'] for item in all_numbers if item['row_type'] == 'id'][:3]
    client_info['teacher_meals'] = [item['number'] for item in all_numbers if item['row_type'] == 'name'][:2]
    return client_info

def export_detailed_client_data_to_dataframe(client_data):
    df_data = []
    for info in client_data:
        row = {'クライアント名': info['client_name'],'園児の給食の数1': info['student_meals'][0] if len(info['student_meals']) > 0 else '','園児の給食の数2': info['student_meals'][1] if len(info['student_meals']) > 1 else '','園児の給食の数3': info['student_meals'][2] if len(info['student_meals']) > 2 else '','先生の給食の数1': info['teacher_meals'][0] if len(info['teacher_meals']) > 0 else '','先生の給食の数2': info['teacher_meals'][1] if len(info['teacher_meals']) > 1 else ''}
        df_data.append(row)
    return pd.DataFrame(df_data)

def get_line_groups(words: List[Dict[str, Any]], y_tolerance: float = 1.2) -> List[List[Dict[str, Any]]]:
    if not words: return []
    sorted_words = sorted(words, key=lambda w: w['top'])
    groups, current_group = [], [sorted_words[0]]
    current_top = sorted_words[0]['top']
    for word in sorted_words[1:]:
        if abs(word['top'] - current_top) <= y_tolerance:
            current_group.append(word)
        else:
            groups.append(current_group)
            current_group, current_top = [word], word['top']
    groups.append(current_group)
    return groups

def get_vertical_boundaries(page, tolerance: float = 2) -> List[float]:
    lines = page.lines
    vertical_lines_x = sorted(list(set(round((line['x0'] + line['x1']) / 2, 1) for line in lines if abs(line['x0'] - line['x1']) < tolerance)))
    words = page.extract_words()
    if not words: return vertical_lines_x
    left_boundary = min(word['x0'] for word in words)
    right_boundary = max(word['x1'] for word in words)
    boundaries = sorted(list(set([round(left_boundary, 1)] + vertical_lines_x + [round(right_boundary, 1)])))
    merged_boundaries = []
    if boundaries:
        merged_boundaries.append(boundaries[0])
        for b in boundaries[1:]:
            if b - merged_boundaries[-1] > tolerance * 2:
                merged_boundaries.append(b)
    return sorted(list(set(merged_boundaries)))

def split_line_using_boundaries(sorted_words_in_line: List[Dict[str, Any]], boundaries: List[float]) -> List[str]:
    columns = [""] * (len(boundaries) - 1)
    for word in sorted_words_in_line:
        word_center_x = (word['x0'] + word['x1']) / 2
        for i in range(len(boundaries) - 1):
            if boundaries[i] <= word_center_x < boundaries[i+1]:
                columns[i] = (columns[i] + " " + word["text"]) if columns[i] else word["text"]
                break
    return columns

def extract_text_with_layout(page) -> List[List[str]]:
    words = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False)
    if not words: return []
    boundaries = get_vertical_boundaries(page)
    if len(boundaries) < 2:
        lines = page.extract_text(layout=False, x_tolerance=3, y_tolerance=3)
        return [[line] for line in lines.split('\n') if line.strip()] if lines else []
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
    for row in rows:
        for c, cell in enumerate(row):
            if c < num_cols and cell.strip(): is_col_empty[c] = False
    keep_indices = [c for c, is_empty in enumerate(is_col_empty) if not is_empty]
    return [[row[i] if i < len(row) else "" for i in keep_indices] for row in rows]

def post_process_rows(rows: List[List[str]]) -> List[List[str]]:
    new_rows = [row[:] for row in rows]
    for i, row in enumerate(new_rows):
        for j, cell in enumerate(row):
            if "合計" in str(cell) and i > 0 and j < len(new_rows[i-1]):
                new_rows[i-1][j] = ""
    return new_rows

def pdf_to_excel_data_for_paste_sheet(pdf_file) -> pd.DataFrame | None:
    try:
        with pdfplumber.open(pdf_file) as pdf:
            if not pdf.pages: return None
            page = pdf.pages[0]
            rows = extract_text_with_layout(page)
            rows = [row for row in rows if any(cell.strip() for cell in row)]
            if not rows: return None
            rows = post_process_rows(rows)
            rows = remove_extra_empty_columns(rows)
            if not rows: return None
            max_cols = max(len(row) for row in rows)
            normalized_rows = [row + [''] * (max_cols - len(row)) for row in rows]
            return pd.DataFrame(normalized_rows)
    except Exception as e:
        st.error(f"PDF処理中にエラーが発生しました: {e}")
        traceback.print_exc()
        return None

def extract_table_from_pdf_for_bento(pdf_file_obj):
    tables = []
    with pdfplumber.open(pdf_file_obj) as pdf:
        for page in pdf.pages:
            if not page.extract_text() or not any(kw in page.extract_text() for kw in ["園名", "飯あり", "キャラ弁"]): continue
            if not page.lines: continue
            table_settings = {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
            table = page.extract_table(table_settings)
            if table: tables.append(table)
    return tables

def find_correct_anchor_for_bento(table, target_row_text="赤"):
    for r_idx, row in enumerate(table):
        if target_row_text in ''.join(str(c) for c in row if c):
            for offset in [1, 2]:
                if r_idx + offset < len(table):
                    for c_idx, cell in enumerate(table[r_idx + offset]):
                        if cell and "飯なし" in cell: return c_idx
    return -1

def extract_bento_range_for_bento(table, start_col):
    bento_list, end_col = [], -1
    for row in table:
        if "おやつ" in ''.join(str(c) for c in row if c):
            for c_idx, cell in enumerate(row):
                if cell and "おやつ" in cell:
                    end_col = c_idx
                    break
            if end_col != -1: break
    if end_col == -1 or start_col >= end_col: return []
    header_row_idx = -1
    for r_idx, row in enumerate(table):
        if any(c and "飯なし" in c for c in row):
            if r_idx > 0: header_row_idx = r_idx - 1
            break
    if header_row_idx == -1: return []
    header_row = table[header_row_idx]
    for col in range(start_col + 1, end_col):
        cell_text = header_row[col] if col < len(header_row) else ""
        if cell_text and str(cell_text).strip():
            bento_list.append(str(cell_text).strip())
    return bento_list

def match_bento_names(pdf_bento_list, master_df):
    if master_df is None or master_df.empty: return [f"{name} (マスタなし)" for name in pdf_bento_list]
    master_tuples = master_df[['商品予定名', 'パン箱入数']].dropna().to_records(index=False).tolist()
    matched = []
    norm_master = [(unicodedata.normalize('NFKC', str(n)).replace(" ", ""), str(n), str(v)) for n, v in master_tuples]
    for pdf_name in pdf_bento_list:
        norm_pdf = unicodedata.normalize('NFKC', str(pdf_name)).replace(" ", "")
        found_match, found_name, found_id = False, None, None
        for norm_m, orig_m, m_id in norm_master:
            if norm_m.startswith(norm_pdf):
                found_name, found_id, found_match = orig_m, m_id, True
                break
        if not found_match:
             for norm_m, orig_m, m_id in norm_master:
                if norm_pdf in norm_m:
                    found_name, found_id, found_match = orig_m, m_id, True
                    break
        if found_match:
            matched.append(f"{found_name} (入数: {found_id})")
        else:
            matched.append(f"{pdf_name} (未マッチ)")
    return matched

# ──────────────────────────────────────────────
# メインアプリケーション
# ──────────────────────────────────────────────
if page_selection == "PDF → Excel 変換":
    st.markdown('<div class="title">【数出表】PDF → Excelへの変換</div>', unsafe_allow_html=True)
    st.markdown("---")
    uploaded_pdf = st.file_uploader("処理するPDFファイルをアップロードしてください", type="pdf")
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
                    ws_paste_n = nouhinsyo_wb["貼り付け用"]
                    for r_idx, row in df_paste_sheet.iterrows():
                        for c_idx, value in enumerate(row):
                            ws_paste_n.cell(row=r_idx + 1, column=c_idx + 1, value=value)
                    if df_bento_for_nouhin is not None: safe_write_df(nouhinsyo_wb["注文弁当の抽出"], df_bento_for_nouhin, start_row=1)
                    if df_client_sheet is not None: safe_write_df(nouhinsyo_wb["クライアント抽出"], df_client_sheet, start_row=1)
                    output_data_only = io.BytesIO()
                    nouhinsyo_wb.save(output_data_only)
                    data_only_excel_bytes = output_data_only.getvalue()
                st.success("✅ ファイルの準備が完了しました！")
                original_pdf_name = os.path.splitext(uploaded_pdf.name)[0]
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(label="【数出表ダウンロード】",data=macro_excel_bytes,file_name=f"{original_pdf_name}_Processed.xlsm",mime="application/vnd.ms-excel.sheet.macroEnabled.12")
                with col2:
                    st.download_button(label="【納品書ダウンロード】",data=data_only_excel_bytes,file_name=f"{original_pdf_name}_Data.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Excelファイル生成中にエラーが発生しました: {e}")
                traceback.print_exc()

elif page_selection == "マスタ設定":
    st.markdown('<div class="title">マスタデータ設定</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">更新するマスタの確認、および新しいCSVファイルのアップロードができます。</div>', unsafe_allow_html=True)

    # ✅ プレビュー表示を「マスタ設定」ページに移動
    st.markdown("##### 現在の商品マスタデータ（プレビュー）")
    if 'master_df' in st.session_state and not st.session_state.master_df.empty:
        st.dataframe(st.session_state.master_df.head(), use_container_width=True)
    else:
        st.warning("商品マスタが読み込まれていません。")
    
    st.markdown("##### 現在の得意先マスタデータ（プレビュー）")
    if 'customer_master_df' in st.session_state and not st.session_state.customer_master_df.empty:
        st.dataframe(st.session_state.customer_master_df.head(), use_container_width=True)
    else:
        st.warning("得意先マスタが読み込まれていません。")
    
    st.markdown("---")

    master_choice = st.selectbox(
        "更新するマスタを選択してください",
        ("商品マスタ", "得意先マスタ")
    )

    if master_choice == "商品マスタ":
        st.markdown("#### 商品マстаの更新")
        master_csv_path = "商品マスタ一覧.csv"
        uploaded_master_csv = st.file_uploader("新しい商品マスタ一覧.csvをアップロード",type="csv",help="ヘッダーには '商品予定名', 'パン箱入数', '商品名' を含めてください。",key="product_master_uploader")
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
                    except Exception: continue
                if new_master_df is not None:
                    st.session_state.master_df = new_master_df
                    new_master_df.to_csv(master_csv_path, index=False, encoding='utf-8-sig')
                    st.success(f"✅ 商品マスタを更新し、'{master_csv_path}' に保存しました。")
                else:
                    st.error("CSVファイルを正しく読み込めませんでした。必須列を確認してください。")
            except Exception as e:
                st.error(f"商品マスタ更新処理中にエラー: {e}")
        st.markdown("##### 現在の商品マスタデータ（全件）")
        if 'master_df' in st.session_state and not st.session_state.master_df.empty:
            st.dataframe(st.session_state.master_df, use_container_width=True)
        else:
            st.warning("商品マスタが読み込まれていません。")

    elif master_choice == "得意先マスタ":
        st.markdown("#### 得意先マスタの更新")
        customer_master_csv_path = "得意先マスタ一覧.csv" # ✅ ファイル名を修正
        uploaded_customer_csv = st.file_uploader("新しい得意先マスタ一覧.csvをアップロード",type="csv",help="ヘッダーには '得意先コード', '得意先名' を含めてください。",key="customer_master_uploader")
        if uploaded_customer_csv is not None:
            try:
                new_customer_df = None
                encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
                for encoding in encodings:
                    try:
                        uploaded_customer_csv.seek(0)
                        temp_df = pd.read_csv(uploaded_customer_csv, encoding=encoding)
                        if all(col in temp_df.columns for col in ['得意先コード', '得意先名']):
                            new_customer_df = temp_df
                            st.info(f"ファイルを {encoding} で読み込みました。")
                            break
                    except Exception: continue
                if new_customer_df is not None:
                    st.session_state.customer_master_df = new_customer_df
                    new_customer_df.to_csv(customer_master_csv_path, index=False, encoding='utf-8-sig')
                    st.success(f"✅ 得意先マスタを更新し、'{customer_master_csv_path}' に保存しました。")
                else:
                    st.error("CSVファイルを正しく読み込めませんでした。必須列を確認してください。")
            except Exception as e:
                st.error(f"得意先マスタ更新処理中にエラー: {e}")
        st.markdown("##### 現在の得意先マスタデータ（全件）")
        if 'customer_master_df' in st.session_state and not st.session_state.customer_master_df.empty:
            st.dataframe(st.session_state.customer_master_df, use_container_width=True)
        else:
            st.warning("得意先マスタが読み込まれていません。")
