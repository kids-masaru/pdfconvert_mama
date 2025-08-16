import streamlit as st
import streamlit.components.v1 as components
import pdfplumber
import pandas as pd
import io
import re
import base64
import os
import unicodedata
import traceback
from typing import List, Dict, Any
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ✅ 修正: st.set_page_config() を最初に移動
st.set_page_config(
    page_title="【数出表】PDF → Excelへの変換",
    page_icon="./static/favicon.ico",
    layout="centered",
)

# --- Streamlit Session Stateの初期化 ---
# （このセクションは変更ありません）
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
        initial_master_df = pd.DataFrame(columns=['商品予定名', 'パン箱入数', '商品名']) # '商品名' を追加
    st.session_state.master_df = initial_master_df


# --- テンプレートExcelファイルの読み込み ---
# template.xlsm の読み込み（変更なし）
if 'template_wb_loaded' not in st.session_state:
    st.session_state.template_wb_loaded = False
    st.session_state.template_wb = None

template_path = "template.xlsm"
if not st.session_state.template_wb_loaded:
    if not os.path.exists(template_path):
        st.error(f"テンプレートファイル '{template_path}' が見つかりません。")
        st.stop()
    try:
        st.session_state.template_wb = load_workbook(template_path, keep_vba=True)
        st.session_state.template_wb_loaded = True
        st.success(f"テンプレートファイル '{template_path}' を読み込みました。")
    except Exception as e:
        st.error(f"テンプレートファイル '{template_path}' の読み込み中にエラーが発生しました: {e}")
        st.stop()

# ✅【変更点】nouhinsyo.xlsx の読み込みを追加
if 'nouhinsyo_wb_loaded' not in st.session_state:
    st.session_state.nouhinsyo_wb_loaded = False
    st.session_state.nouhinsyo_wb = None

nouhinsyo_path = "nouhinsyo.xlsx"
if not st.session_state.nouhinsyo_wb_loaded:
    if not os.path.exists(nouhinsyo_path):
        st.error(f"納品書ファイル '{nouhinsyo_path}' が見つかりません。")
        st.stop()
    try:
        st.session_state.nouhinsyo_wb = load_workbook(nouhinsyo_path)
        st.session_state.nouhinsyo_wb_loaded = True
        st.success(f"納品書ファイル '{nouhinsyo_path}' を読み込みました。")
    except Exception as e:
        st.error(f"納品書ファイル '{nouhinsyo_path}' の読み込み中にエラーが発生しました: {e}")
        st.stop()


# --- HTML/CSS, サイドバー, 関数定義 ---
# (このセクションは変更ありません。長いので省略します)
# PWA用HTML埋め込み
components.html(
    """
    <link rel="manifest" href="./static/manifest.json">
    <link rel="icon" href="./static/favicon.ico">
    <link rel="apple-touch-icon" sizes="180x180" href="./static/icons/apple-touch-icon.png">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-title" content="PDFConverter">
    """,
    height=0,
)
# CSSスタイル
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Roboto:wght@300;400;500&display=swap');
        .stApp { background: #fff5e6; font-family: 'Inter', sans-serif; }
        .title { font-size: 1.5rem; font-weight: 600; color: #333; margin-bottom: 5px; }
        .subtitle { font-size: 0.9rem; color: #666; margin-bottom: 25px; }
    </style>
""", unsafe_allow_html=True)
# サイドバーナビゲーション
st.sidebar.title("メニュー")
page_selection = st.sidebar.radio("表示する機能を選択してください", ("PDF → Excel 変換", "マスタ設定"), index=0)
st.markdown("---")
# 関数定義... (変更なし)
def extract_detailed_client_info_from_pdf(pdf_file_obj):
    """PDFから詳細なクライアント情報（名前＋給食の数）を抽出する"""
    client_data = []

    try:
        with pdfplumber.open(pdf_file_obj) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # 表形式のデータを抽出
                rows = extract_text_with_layout(page)
                if not rows:
                    continue

                # 園名の位置を探す
                garden_row_idx = -1
                for i, row in enumerate(rows):
                    row_text = ''.join(str(cell) for cell in row if cell)
                    if '園名' in row_text:
                        garden_row_idx = i
                        break

                if garden_row_idx == -1:
                    continue

                # 園名より下の行を処理
                current_client_id = None
                current_client_name = None

                for i in range(garden_row_idx + 1, len(rows)):
                    row = rows[i]

                    # 10001が出てきたら終了
                    row_text = ''.join(str(cell) for cell in row if cell)
                    if '10001' in row_text:
                        break

                    # 空行はスキップ
                    if not any(str(cell).strip() for cell in row):
                        continue

                    # 左の列（1番目の列）をチェック
                    if len(row) > 0 and row[0]:
                        left_cell = str(row[0]).strip()

                        # 数字だけの場合はID
                        if re.match(r'^\d+$', left_cell):
                            # 前のクライアントのデータを保存
                            if current_client_id and current_client_name:
                                client_info = extract_meal_numbers_from_row(rows, i-1, current_client_id, current_client_name)
                                if client_info:
                                    client_data.append(client_info)

                            current_client_id = left_cell
                            current_client_name = None

                        # 数字以外の場合はクライアント名
                        elif not re.match(r'^\d+$', left_cell) and current_client_id:
                            current_client_name = left_cell

                # 最後のクライアントのデータを保存
                if current_client_id and current_client_name:
                    client_info = extract_meal_numbers_from_row(rows, len(rows)-1, current_client_id, current_client_name)
                    if client_info:
                        client_data.append(client_info)

    except Exception as e:
        st.error(f"クライアント情報抽出中にエラーが発生しました: {e}")

    return client_data

def extract_meal_numbers_from_row(rows, row_idx, client_id, client_name):
    """指定された行とその周辺から給食の数を抽出"""
    client_info = {
        'client_id': client_id,
        'client_name': client_name,
        'student_meals': [],
        'teacher_meals': []
    }

    # IDの行とクライアント名の行から数字を抽出
    rows_to_check = []

    # IDの行を探す
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

    # 数字を抽出
    all_numbers = []

    for row_type, idx, row in rows_to_check:
        # 左の列（0番目）以外の列から数字を抽出
        for col_idx in range(1, len(row)):
            cell = str(row[col_idx]).strip()
            if cell and re.match(r'^\d+$', cell):
                all_numbers.append({
                    'number': int(cell),
                    'row_type': row_type,
                    'col_idx': col_idx
                })
            elif cell and not re.match(r'^\d+$', cell) and cell != '':
                # 数字以外の文字が出てきたらその行はここで終了
                break

    # 園児の給食の数と先生の給食の数に分ける
    # IDの行の数字は園児の給食の数
    # クライアント名の行の数字は先生の給食の数

    id_numbers = [item['number'] for item in all_numbers if item['row_type'] == 'id']
    name_numbers = [item['number'] for item in all_numbers if item['row_type'] == 'name']

    # 園児の給食の数（最大3つ）
    client_info['student_meals'] = id_numbers[:3]

    # 先生の給食の数（最大2つ）
    client_info['teacher_meals'] = name_numbers[:2]

    return client_info

def export_detailed_client_data_to_dataframe(client_data):
    """詳細クライアント情報をDataFrameに変換"""
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

def is_number(text: str) -> bool:
    return bool(re.match(r'^\d+$', text.strip()))

def get_line_groups(words: List[Dict[str, Any]], y_tolerance: float = 1.2) -> List[List[Dict[str, Any]]]:
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: w['top'])
    groups = []
    current_group = [sorted_words[0]]
    current_top = sorted_words[0]['top']
    for word in sorted_words[1:]:
        if abs(word['top'] - current_top) <= y_tolerance:
            current_group.append(word)
        else:
            groups.append(current_group)
            current_group = [word]
            current_top = word['top']
    groups.append(current_group)
    return groups

def get_vertical_boundaries(page, tolerance: float = 2) -> List[float]:
    vertical_lines_x = []
    for line in page.lines:
        if abs(line['x0'] - line['x1']) < tolerance:
            vertical_lines_x.append((line['x0'] + line['x1']) / 2)
    vertical_lines_x = sorted(list(set(round(x, 1) for x in vertical_lines_x)))

    words = page.extract_words()
    if not words:
        return vertical_lines_x

    left_boundary = min(word['x0'] for word in words)
    right_boundary = max(word['x1'] for word in words)

    boundaries = sorted(list(set([round(left_boundary, 1)] + vertical_lines_x + [round(right_boundary, 1)])))

    merged_boundaries = []
    if boundaries:
        merged_boundaries.append(boundaries[0])
        for i in range(1, len(boundaries)):
            if boundaries[i] - merged_boundaries[-1] > tolerance * 2:
                merged_boundaries.append(boundaries[i])
        if right_boundary > merged_boundaries[-1] + tolerance * 2:
            merged_boundaries.append(round(right_boundary, 1))
        boundaries = sorted(list(set(merged_boundaries)))

    return boundaries

def split_line_using_boundaries(sorted_words_in_line: List[Dict[str, Any]], boundaries: List[float]) -> List[str]:
    columns = [""] * (len(boundaries) - 1)
    for word in sorted_words_in_line:
        word_center_x = (word['x0'] + word['x1']) / 2
        for i in range(len(boundaries) - 1):
            left = boundaries[i]
            right = boundaries[i + 1]
            if left <= word_center_x < right:
                if columns[i]:
                    columns[i] += " " + word["text"]
                else:
                    columns[i] = word["text"]
                break
    return columns

def extract_text_with_layout(page) -> List[List[str]]:
    words = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False)
    if not words:
        return []

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
    if not rows:
        return rows
    num_cols = max(len(row) for row in rows) if rows else 0
    if num_cols == 0:
        return rows

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

def pdf_to_excel_data_for_paste_sheet(pdf_file) -> pd.DataFrame | None:
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
        traceback.print_exc()
        return None

def extract_table_from_pdf_for_bento(pdf_file_obj):
    tables = []
    with pdfplumber.open(pdf_file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text is None: continue

            start_keywords = ["園名", "飯あり", "キャラ弁"]
            end_keywords = ["おやつ", "合計", "PAGE"]

            if not any(kw in text for kw in start_keywords):
                continue

            lines = page.lines
            if not lines:
                continue

            y_coords = sorted(set([line['top'] for line in lines] + [line['bottom'] for line in lines]))
            if len(y_coords) < 2:
                continue

            table_top = min(y_coords)
            table_bottom = max(y_coords)

            x_coords = sorted(set([line['x0'] for line in lines] + [line['x1'] for line in lines]))
            if len(x_coords) < 2:
                continue

            table_left = min(x_coords)
            table_right = max(x_coords)

            table_bbox = (table_left, table_top, table_right, table_bottom)
            cropped_page = page.crop(table_bbox)

            table_settings = {
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance": 3,
                "join_tolerance": 3,
                "edge_min_length": 15,
            }

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

    if anchor_row_idx == -1:
        return []

    if anchor_row_idx - 1 >= 0:
        header_row_idx = anchor_row_idx - 1
    else:
        return []

    for col in range(start_col + 1, end_col + 1):
        if col < len(table[header_row_idx]):
            cell_text = table[header_row_idx][col]
        else:
            cell_text = ""

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
        elif '商品予定名' in master_df.columns:
            st.warning("マスタデータに「パン箱入数」列が見つかりません。")
            master_data_tuples = master_df['商品予定名'].dropna().astype(str).tolist()
            master_data_tuples = [(name, "") for name in master_data_tuples]
        else:
            st.error("マスタデータに「商品予定名」列が見つかりません。")
            return [f"{name} (商品予定名列なし)" for name in pdf_bento_list]

    except Exception as e:
        st.error(f"マスタデータ処理中にエラーが発生しました: {e}")
        return [f"{name} (処理エラー)" for name in pdf_bento_list]

    if len(master_data_tuples) == 0:
        st.warning("マスタデータから有効な商品情報が抽出できませんでした。")
        return [f"{name} (マスタ空)" for name in pdf_bento_list]

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

        if not found_match:
            for num_chars_to_remove in range(1, 4):
                if len(original_normalized_pdf_name) > num_chars_to_remove:
                    truncated_pdf_name = original_normalized_pdf_name[:-num_chars_to_remove]

                    for norm_m_name, orig_m_name, m_id in normalized_master_data_tuples:
                        if norm_m_name.startswith(truncated_pdf_name):
                            found_original_master_name = orig_m_name
                            found_id = m_id
                            found_match = True
                            break

                    if not found_match:
                        for norm_m_name, orig_m_name, m_id in normalized_master_data_tuples:
                            if truncated_pdf_name in norm_m_name:
                                found_original_master_name = orig_m_name
                                found_id = m_id
                                found_match = True
                                break

                    if found_match:
                        break

        if found_original_master_name:
            if found_id:
                matched.append(f"{found_original_master_name} (入数: {found_id})")
            else:
                matched.append(found_original_master_name)
        else:
            matched.append(f"{pdf_name} (未マッチ)")

    return matched


# ──────────────────────────────────────────────
# メインアプリケーション
# ──────────────────────────────────────────────

# PDF → Excel 変換 ページ
if page_selection == "PDF → Excel 変換":
    st.markdown('<div class="title">【数出表】PDF → Excelへの変換</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">PDFの数出表をExcelに変換し、詳細なクライアント情報も含めて一括処理します。</div>', unsafe_allow_html=True)

    uploaded_pdf = st.file_uploader("処理するPDFファイルをアップロードしてください", type="pdf")

    if uploaded_pdf is not None and st.session_state.template_wb is not None and st.session_state.nouhinsyo_wb is not None:
        # PDFのバイナリデータをio.BytesIOに変換
        pdf_bytes_io = io.BytesIO(uploaded_pdf.getvalue())

        # 1. 貼り付け用データの抽出
        df_paste_sheet = None
        with st.spinner("「貼り付け用」データを抽出中..."):
            pdf_bytes_io.seek(0)
            df_paste_sheet = pdf_to_excel_data_for_paste_sheet(pdf_bytes_io)

        # 2. 注文弁当データの抽出
        df_bento_sheet = None
        if df_paste_sheet is not None:
            with st.spinner("「注文弁当の抽出」データを抽出中..."):
                try:
                    pdf_bytes_io.seek(0)
                    tables = extract_table_from_pdf_for_bento(pdf_bytes_io)
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
                                            bento_name = item[:match.start()]
                                            bento_count = match.group(1)
                                            output_data_bento.append([bento_name.strip(), bento_count.strip()])
                                        elif "(未マッチ)" in item:
                                            bento_name = item.replace(" (未マッチ)", "").strip()
                                            output_data_bento.append([bento_name, ""])
                                        else:
                                            output_data_bento.append([item.strip(), ""])
                                    df_bento_sheet = pd.DataFrame(output_data_bento, columns=['商品予定名', 'パン箱入数'])
                except Exception as e:
                    st.error(f"注文弁当データ処理中にエラー: {e}")
                    traceback.print_exc()

        # 3. 詳細クライアント情報の抽出
        df_client_sheet = None
        if df_paste_sheet is not None:
            with st.spinner("「クライアント抽出」データを抽出中..."):
                try:
                    pdf_bytes_io.seek(0)
                    client_data = extract_detailed_client_info_from_pdf(pdf_bytes_io)
                    if client_data:
                        df_client_sheet = export_detailed_client_data_to_dataframe(client_data)
                        st.success(f"クライアント情報 {len(client_data)} 件を抽出しました")
                    else:
                        st.warning("クライアント情報を抽出できませんでした。")
                except Exception as e:
                    st.error(f"クライアント情報抽出中にエラー: {e}")
                    traceback.print_exc()

        # 4. Excelファイルへの書き込みと生成
        if df_paste_sheet is not None:
            try:
                # --- template.xlsmへの書き込み ---
                with st.spinner("template.xlsm にデータを書き込み中..."):
                    # 貼り付け用
                    try:
                        ws_paste = st.session_state.template_wb["貼り付け用"]
                        for r_idx, row in df_paste_sheet.iterrows():
                            for c_idx, value in enumerate(row):
                                ws_paste.cell(row=r_idx + 1, column=c_idx + 1, value=value)
                    except KeyError:
                        st.error("template.xlsmに「貼り付け用」シートが見つかりません。")
                        st.stop()
                    # 注文弁当
                    if df_bento_sheet is not None and not df_bento_sheet.empty:
                        try:
                            ws_bento = st.session_state.template_wb["注文弁当の抽出"]
                            for r in dataframe_to_rows(df_bento_sheet, index=False, header=True):
                                ws_bento.append(r)
                        except KeyError:
                            st.error("template.xlsmに「注文弁当の抽出」シートが見つかりません。")
                    # クライアント抽出
                    if df_client_sheet is not None and not df_client_sheet.empty:
                        try:
                            ws_client = st.session_state.template_wb["クライアント抽出"]
                            for r in dataframe_to_rows(df_client_sheet, index=False, header=True):
                                ws_client.append(r)
                        except KeyError:
                            st.error("template.xlsmに「クライアント抽出」シートが見つかりません。")
                
                # --- nouhinsyo.xlsxへの書き込み ---
                with st.spinner("nouhinsyo.xlsx にデータを書き込み中..."):
                    # ✅【変更点】nouhinsyo.xlsx用の弁当データフレームを作成
                    df_bento_for_nouhin = None
                    if df_bento_sheet is not None and not df_bento_sheet.empty:
                        master_df = st.session_state.master_df
                        # マスターデータから「商品予定名」をキーに「商品名」を引くための辞書を作成
                        # `drop_duplicates` を追加して、キーの重複エラーを防ぐ
                        master_map = master_df.drop_duplicates(subset=['商品予定名']).set_index('商品予定名')['商品名'].to_dict()
                        
                        df_bento_for_nouhin = df_bento_sheet.copy()
                        # map関数を使って「商品名」列を追加
                        df_bento_for_nouhin['商品名'] = df_bento_for_nouhin['商品予定名'].map(master_map)
                        # 列の順番を A:商品予定名, B:パン箱入数, C:商品名 にする
                        df_bento_for_nouhin = df_bento_for_nouhin[['商品予定名', 'パン箱入数', '商品名']]

                    # 貼り付け用
                    try:
                        ws_paste_n = st.session_state.nouhinsyo_wb["貼り付け用"]
                        for r_idx, row in df_paste_sheet.iterrows():
                            for c_idx, value in enumerate(row):
                                ws_paste_n.cell(row=r_idx + 1, column=c_idx + 1, value=value)
                    except KeyError:
                        st.error("nouhinsyo.xlsxに「貼り付け用」シートが見つかりません。")
                    # 注文弁当
                    if df_bento_for_nouhin is not None and not df_bento_for_nouhin.empty:
                        try:
                            ws_bento_n = st.session_state.nouhinsyo_wb["注文弁当の抽出"]
                            for r in dataframe_to_rows(df_bento_for_nouhin, index=False, header=True):
                                ws_bento_n.append(r)
                        except KeyError:
                            st.error("nouhinsyo.xlsxに「注文弁当の抽出」シートが見つかりません。")
                    # クライアント抽出
                    if df_client_sheet is not None and not df_client_sheet.empty:
                        try:
                            ws_client_n = st.session_state.nouhinsyo_wb["クライアント抽出"]
                            for r in dataframe_to_rows(df_client_sheet, index=False, header=True):
                                ws_client_n.append(r)
                        except KeyError:
                            st.error("nouhinsyo.xlsxに「クライアント抽出」シートが見つかりません。")


                # --- ファイルのバイナリ生成 ---
                # template.xlsm
                output_macro = io.BytesIO()
                st.session_state.template_wb.save(output_macro)
                output_macro.seek(0)
                macro_excel_bytes = output_macro.read()
                
                # nouhinsyo.xlsx
                output_data_only = io.BytesIO()
                st.session_state.nouhinsyo_wb.save(output_data_only)
                output_data_only.seek(0)
                data_only_excel_bytes = output_data_only.read()

                # --- 処理完了とダウンロード ---
                st.success("✅ 全ての処理が完了しました！")
                
                col1, col2 = st.columns(2)

                with col1:
                    st.download_button(
                        label="📥 template.xlsmをダウンロード",
                        data=macro_excel_bytes,
                        file_name="template.xlsm", # ✅【変更点】ファイル名を固定
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    )

                with col2:
                    st.download_button(
                        label="📥 nouhinsyo.xlsxをダウンロード",
                        data=data_only_excel_bytes,
                        file_name="nouhinsyo.xlsx", # ✅【変更点】ファイル名を固定
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

            except Exception as e:
                st.error(f"Excelファイル生成中にエラーが発生しました: {e}")
                traceback.print_exc()

# マスタ設定 ページ
elif page_selection == "マスタ設定":
    # (このセクションは変更ありません)
    st.markdown('<div class="title">マスタデータ設定</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">商品マスタのCSVファイルをアップロードして更新します。</div>', unsafe_allow_html=True)

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
                except UnicodeDecodeError:
                    continue
            
            if new_master_df is not None:
                st.session_state.master_df = new_master_df
                # ローカルファイルに保存
                try:
                    new_master_df.to_csv(master_csv_path, index=False, encoding='utf-8-sig')
                    st.success("マスタデータを更新し、ローカルファイルに保存しました。")
                except Exception as e:
                    st.error(f"マスタデータのローカル保存中にエラーが発生しました: {e}")
            else:
                st.error("アップロードされたファイルに必要なヘッダー（'商品予定名', 'パン箱入数', '商品名'）が見つかりませんでした。")
        except Exception as e:
            st.error(f"マスタCSVの読み込み中にエラーが発生しました: {e}")
            traceback.print_exc()

    st.markdown("---")
    st.markdown("#### 現在のマスタデータ")
    if 'master_df' in st.session_state and st.session_state.master_df is not None:
        st.dataframe(st.session_state.master_df)
        csv = st.session_state.master_df.to_csv(index=False, encoding='utf-8-sig')
        b64 = base64.b64encode(csv.encode()).decode()
        href = f'<a href="data:file/csv;base64,{b64}" download="商品マスタ一覧.csv">現在のマスタをダウンロード</a>'
        st.markdown(href, unsafe_allow_html=True)
    else:
        st.info("現在、読み込まれているマスタデータはありません。")
