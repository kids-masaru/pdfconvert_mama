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

# ✅ 修正: st.set_page_config() を最初に移動
st.set_page_config(
    page_title="【数出表】PDF → Excelへの変換",
    page_icon="./static/favicon.ico", # faviconのパスを修正
    layout="centered",
)

# --- Streamlit Session Stateの初期化 ---
# マスタデータをセッションステートで管理し、アプリ実行中に保持する
if 'master_df' not in st.session_state:
    # アプリ起動時に既存の商品マスタCSVを読み込む試み
    master_csv_path = "商品マスタ一覧.csv"
    initial_master_df = None
    if os.path.exists(master_csv_path):
        # ✅ 読み込みエンコーディングに utf-8-sig を追加
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
        initial_master_df = pd.DataFrame(columns=['商品予定名', 'パン箱入数']) # 空のDataFrameで初期化
    st.session_state.master_df = initial_master_df

# テンプレートExcelファイルのパス設定と存在確認 (セッションステートで管理)
if 'template_wb_loaded' not in st.session_state:
    st.session_state.template_wb_loaded = False
    st.session_state.template_wb = None

template_path = "template.xlsm"

if not st.session_state.template_wb_loaded:
    if not os.path.exists(template_path):
        st.error(f"テンプレートファイル '{template_path}' が見つかりません。スクリプトと同じ場所に配置してください。")
        st.stop()
    
    try:
        st.session_state.template_wb = load_workbook(template_path, keep_vba=True)
        st.session_state.template_wb_loaded = True
        st.success(f"テンプレートファイル '{template_path}' を読み込みました。")
    except Exception as e:
        st.error(f"テンプレートファイル '{template_path}' の読み込み中にエラーが発生しました: {e}")
        st.session_state.template_wb = None
        st.stop()

# ──────────────────────────────────────────────
# ① HTML <head> 埋め込み（PWA用 manifest & 各種アイコン）
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
# ③ CSS／UI スタイル定義
# ──────────────────────────────────────────────
st.markdown("""
    <style>
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
    </style>
""", unsafe_allow_html=True)


# --- サイドバーナビゲーション ---
st.sidebar.title("メニュー")
page_selection = st.sidebar.radio(
    "表示する機能を選択してください",
    ("PDF → Excel 変換", "クライアント情報抽出", "マスタ設定"),
    index=0 # 初期表示は「PDF → Excel 変換」
)

st.markdown("---") # メインコンテンツとサイドバーの区切り


# --- メインコンテンツの表示ロジック ---

# メインコンテナ開始
st.markdown('<div class="main-container">', unsafe_allow_html=True)

# ──────────────────────────────────────────────
# クライアント情報抽出の新しい関数群
# ──────────────────────────────────────────────

def extract_client_and_meal_info_from_pdf(pdf_file_obj):
    """PDFから園名の下のクライアント名と給食の数を抽出する"""
    client_data = []
    
    try:
        with pdfplumber.open(pdf_file_obj) as pdf:
            for page_num, page in enumerate(pdf.pages):
                st.write(f"📄 ページ {page_num + 1} を処理中...")
                
                # 縦線の位置を検出
                vertical_lines = []
                if page.lines:
                    for line in page.lines:
                        # 縦線を検出（x0とx1がほぼ同じで、y0とy1が異なる）
                        if abs(line['x0'] - line['x1']) < 2:  # 縦線の判定
                            vertical_lines.append(line['x0'])
                
                # 縦線をソート
                vertical_lines = sorted(list(set(vertical_lines)))
                st.write(f"📏 検出された縦線の位置: {vertical_lines}")
                
                # テキストを抽出
                text = page.extract_text()
                if not text:
                    st.write("❌ このページからテキストを抽出できませんでした")
                    continue
                
                # デバッグ：抽出されたテキストの一部を表示
                st.write(f"🔍 抽出されたテキストの最初の300文字:")
                st.code(text[:300])
                
                lines = text.split('\n')
                st.write(f"📝 総行数: {len(lines)}")
                
                # 園名を探す
                garden_found = False
                start_index = -1
                for i, line in enumerate(lines):
                    if '園名' in line:
                        garden_found = True
                        start_index = i + 1
                        st.write(f"✅ 園名を発見！ 行番号: {i}, 内容: '{line}'")
                        break
                
                if not garden_found:
                    st.write("❌ 園名が見つかりませんでした")
                    continue
                
                # 園名の下からクライアント名と給食情報を抽出
                st.write(f"🔍 園名の下の行から抽出開始（行 {start_index} から）:")
                
                i = start_index
                while i < len(lines):
                    line = lines[i].strip()
                    
                    # 10000が出てきたら終了
                    if '10000' in line:
                        st.write(f"🛑 10000を発見！ 処理終了")
                        break
                    
                    # 空行はスキップ
                    if not line:
                        i += 1
                        continue
                    
                    # IDで始まる行を検出
                    if re.match(r'^\d+', line):
                        client_info = extract_client_info_from_lines(lines, i)
                        if client_info:
                            client_data.append(client_info)
                            st.write(f"✅ 抽出完了: {client_info}")
                        i += client_info.get('lines_processed', 1)
                    else:
                        i += 1
    
    except Exception as e:
        st.error(f"クライアント情報抽出中にエラーが発生しました: {e}")
        st.exception(e)
    
    # 結果の表示
    st.write(f"🎯 最終結果: 総 {len(client_data)} 件のクライアント情報を抽出")
    if client_data:
        st.write("抽出されたクライアント情報:")
        for i, data in enumerate(client_data[:5]):  # 最初の5件を表示
            st.write(f"  {i+1}: {format_client_info(data)}")
        if len(client_data) > 5:
            st.write(f"  ... 他 {len(client_data) - 5} 件")
    
    return client_data


def extract_client_info_from_lines(lines, start_idx):
    """指定された行から1つのクライアント情報を抽出"""
    client_info = {
        'client_name': '',
        'student_meals': [],
        'teacher_meals': [],
        'lines_processed': 1
    }
    
    current_line = lines[start_idx].strip()
    
    # IDを抽出
    id_match = re.match(r'^(\d+)', current_line)
    if not id_match:
        return None
    
    client_id = id_match.group(1)
    remaining_text = current_line[len(client_id):].strip()
    
    # 同じ行にクライアント名があるかチェック
    if remaining_text and not remaining_text.isdigit():
        # 同じ行にクライアント名がある場合
        client_info['client_name'] = extract_client_name_from_text(remaining_text)
        # 残りの部分から園児の給食の数を抽出
        client_info['student_meals'] = extract_numbers_from_text(remaining_text, after_name=True)
        client_info['lines_processed'] = 1
    else:
        # 次の行にクライアント名がある場合
        if start_idx + 1 < len(lines):
            next_line = lines[start_idx + 1].strip()
            if next_line and not next_line.isdigit():
                client_info['client_name'] = extract_client_name_from_text(next_line)
                client_info['teacher_meals'] = extract_numbers_from_text(next_line, after_name=True)
                client_info['lines_processed'] = 2
        
        # IDの行から園児の給食の数を抽出
        if remaining_text:
            client_info['student_meals'] = extract_numbers_from_text(remaining_text)
    
    # 追加の行もチェック（園児の給食の数や先生の給食の数のため）
    check_lines = 3  # 最大3行先まで確認
    for offset in range(1, min(check_lines + 1, len(lines) - start_idx)):
        check_line = lines[start_idx + offset].strip()
        
        # 数字以外の文字が出てきたら終了
        if check_line and not check_line.replace(' ', '').isdigit() and not has_numbers(check_line):
            break
        
        # 数字が含まれる行から追加の給食の数を抽出
        if has_numbers(check_line):
            additional_numbers = extract_numbers_from_text(check_line)
            
            # 園児の給食の数が足りない場合は園児に追加
            if len(client_info['student_meals']) < 3:
                needed = 3 - len(client_info['student_meals'])
                client_info['student_meals'].extend(additional_numbers[:needed])
                additional_numbers = additional_numbers[needed:]
            
            # 残りは先生の給食の数に追加
            client_info['teacher_meals'].extend(additional_numbers)
            
            if offset >= client_info['lines_processed']:
                client_info['lines_processed'] = offset + 1
    
    return client_info


def extract_client_name_from_text(text):
    """テキストからクライアント名を抽出（数字以外の部分）"""
    # 数字を除去してクライアント名を抽出
    name_parts = []
    words = text.split()
    
    for word in words:
        if not word.isdigit():
            name_parts.append(word)
        else:
            # 数字が出てきたら名前の部分は終了
            break
    
    return ' '.join(name_parts).strip()


def extract_numbers_from_text(text, after_name=False):
    """テキストから数字を抽出"""
    numbers = []
    
    if after_name:
        # 名前の後の数字を抽出
        words = text.split()
        name_ended = False
        
        for word in words:
            if word.isdigit():
                numbers.append(int(word))
                name_ended = True
            elif name_ended:
                # 名前が終わった後に数字以外が出てきたら終了
                break
    else:
        # 単純に数字を抽出
        words = text.split()
        for word in words:
            if word.isdigit():
                numbers.append(int(word))
    
    return numbers


def has_numbers(text):
    """テキストに数字が含まれているかチェック"""
    return any(char.isdigit() for char in text)


def format_client_info(client_info):
    """クライアント情報を見やすい形式でフォーマット"""
    name = client_info['client_name']
    student_meals = client_info['student_meals']
    teacher_meals = client_info['teacher_meals']
    
    # 園児の給食の数を3つまで表示（足りない場合は空白）
    student_meal_str = []
    for i in range(3):
        if i < len(student_meals):
            student_meal_str.append(str(student_meals[i]))
        else:
            student_meal_str.append('')
    
    # 先生の給食の数を2つまで表示（足りない場合は空白）
    teacher_meal_str = []
    for i in range(2):
        if i < len(teacher_meals):
            teacher_meal_str.append(str(teacher_meals[i]))
        else:
            teacher_meal_str.append('')
    
    return f"{name}\t{student_meal_str[0]}\t{student_meal_str[1]}\t{student_meal_str[2]}\t{teacher_meal_str[0]}\t{teacher_meal_str[1]}"


def export_client_data_to_csv(client_data, filename="client_meal_data.csv"):
    """クライアント情報をCSVファイルに出力"""
    # データフレーム用のリストを作成
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
    
    # データフレームを作成してCSVに出力
    df = pd.DataFrame(df_data)
    
    return df


# ──────────────────────────────────────────────
# 既存のPDF→Excel変換ロジック
# ──────────────────────────────────────────────

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
        if right_boundary > merged_boundaries[-1] + tolerance * 2 :
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
                st.warning("PDFの最初のページからテキストデータを抽出できませんでした。（貼り付け用）")
                return None

            rows = post_process_rows(rows)
            rows = remove_extra_empty_columns(rows)
            if not rows or not rows[0]:
                    st.warning("空の列を削除した結果、データがなくなりました。（貼り付け用）")
                    return None

            max_cols = max(len(row) for row in rows) if rows else 0
            normalized_rows = [row + [''] * (max_cols - len(row)) for row in rows]
            df = pd.DataFrame(normalized_rows)
            return df

    except Exception as e:
        st.error(f"PDF処理中にエラーが発生しました（貼り付け用）: {e}")
        return None

def extract_table_from_pdf_for_bento(pdf_file_obj):
    tables = []
    with pdfplumber.open(pdf_file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            
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
        st.error("マスタデータがロードされていません。マスタ設定ページでCSVをアップロードしてください。")
        return [f"{name} (マスタデータなし)" for name in pdf_bento_list]

    master_data_tuples = []
    try:
        if '商品予定名' in master_df.columns and 'パン箱入数' in master_df.columns:
            master_data_tuples = master_df[['商品予定名', 'パン箱入数']].dropna().values.tolist()
            master_data_tuples = [(str(name), str(value)) for name, value in master_data_tuples]
        elif '商品予定名' in master_df.columns:
            st.warning("警告: マスタデータに「パン箱入数」列が見つかりません。商品予定名のみで照合します。")
            master_data_tuples = master_df['商品予定名'].dropna().astype(str).tolist()
            master_data_tuples = [(name, "") for name in master_data_tuples]
        else:
            st.error("エラー: マスタデータに「商品予定名」列が見つかりません。")
            return [f"{name} (商品予定名列なし)" for name in pdf_bento_list]

    except KeyError as e:
        st.error(f"エラー: マスタデータに必要な列が見つかりません: {e}。CSVのヘッダー名を確認してください。")
        return [f"{name} (列エラー)" for name in pdf_bento_list]
    except Exception as e:
        st.error(f"マスタデータ処理中に予期せぬエラーが発生しました: {e}")
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
        
        current_pdf_name_for_matching = original_normalized_pdf_name
        
        found_match = False
        found_original_master_name = None
        found_id = None
        
        for norm_m_name, orig_m_name, m_id in normalized_master_data_tuples:
            if norm_m_name.startswith(current_pdf_name_for_matching):
                found_original_master_name = orig_m_name
                found_id = m_id
                found_match = True
                break
        
        if not found_match:
            for norm_m_name, orig_m_name, m_id in normalized_master_data_tuples:
                if current_pdf_name_for_matching in norm_m_name:
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

def extract_client_names_from_pdf(pdf_file_obj):
    """PDFから園名の下のクライアント名を抽出する（縦線を考慮した改良版）"""
    client_names = []

    try:
        with pdfplumber.open(pdf_file_obj) as pdf:
            for page_num, page in enumerate(pdf.pages):
                st.write(f"📄 ページ {page_num + 1} を処理中...")
            
                # 縦線の位置を検出
                vertical_lines = []
                if page.lines:
                    for line in page.lines:
                        # 縦線を検出（x0とx1がほぼ同じで、y0とy1が異なる）
                        if abs(line['x0'] - line['x1']) < 2:  # 縦線の判定
                            vertical_lines.append(line['x0'])
            
                # 縦線をソート
                vertical_lines = sorted(list(set(vertical_lines)))
                st.write(f"📏 検出された縦線の位置: {vertical_lines}")
            
                # 1つ目と2つ目の縦線の間の範囲を設定
                if len(vertical_lines) >= 2:
                    left_bound = vertical_lines[0]
                    right_bound = vertical_lines[1]
                    st.write(f"🎯 クライアント名抽出範囲: x={left_bound:.1f} から x={right_bound:.1f}")
                else:
                    st.write("⚠️ 十分な縦線が検出されませんでした。テキスト全体から抽出します。")
                    left_bound = None
                    right_bound = None
            
                # テキストを抽出（縦線の範囲を考慮）
                if left_bound is not None and right_bound is not None:
                    # 指定された範囲内のテキストを抽出
                    crop_box = (left_bound, 0, right_bound, page.height)
                    cropped_page = page.crop(crop_box)
                    text = cropped_page.extract_text()
                    st.write(f"📄 範囲指定でのテキスト抽出完了")
                else:
                    # 通常のテキスト抽出
                    text = page.extract_text()
            
                if not text:
                    st.write("❌ このページからテキストを抽出できませんでした")
                    continue
            
                # デバッグ：抽出されたテキストの一部を表示
                st.write(f"🔍 抽出されたテキストの最初の300文字:")
                st.code(text[:300])
            
                lines = text.split('\n')
                st.write(f"📝 総行数: {len(lines)}")
            
                # 園名を探す
                garden_found = False
                start_index = -1
                for i, line in enumerate(lines):
                    if '園名' in line:
                        garden_found = True
                        start_index = i + 1
                        st.write(f"✅ 園名を発見！ 行番号: {i}, 内容: '{line}'")
                        break
            
                if not garden_found:
                    st.write("❌ 園名が見つかりませんでした")
                    continue
            
                # 園名の下からクライアント名を抽出
                st.write(f"🔍 園名の下の行から抽出開始（行 {start_index} から）:")
            
                extracted_count = 0
                expect_id = True  # 最初はIDを期待
            
                for j in range(start_index, len(lines)):
                    line = lines[j].strip()
                
                    # 10000が出てきたら終了
                    if '10000' in line:
                        st.write(f"🛑 10000を発見！ 処理終了")
                        break
                
                    # 空行はスキップ
                    if not line:
                        continue
                
                    # IDとクライアント名の交互パターンを処理
                    if expect_id:
                        # IDを期待している場合
                        if re.match(r'^\d+', line):
                            # IDが見つかった
                            id_match = re.match(r'^(\d+)', line)
                            if id_match:
                                current_id = id_match.group(1)
                                st.write(f"🔢 ID発見: '{current_id}'")
                                expect_id = False  # 次はクライアント名を期待
                            
                                # 同じ行にクライアント名が含まれているかチェック
                                remaining_text = line[len(current_id):].strip()
                                if remaining_text and not remaining_text.isdigit():
                                    # 同じ行にクライアント名がある
                                    client_names.append(remaining_text)
                                    extracted_count += 1
                                    st.write(f"✅ 抽出（ID+名前）: '{remaining_text}'")
                                    expect_id = True  # 次はまたIDを期待
                        else:
                            # IDが期待されているのに見つからない場合はスキップ
                            st.write(f"⏭️ ID期待中だが非数字行をスキップ: '{line}'")
                    else:
                        # クライアント名を期待している場合
                        if not line.isdigit():
                            # 数字以外の行をクライアント名として抽出
                            client_names.append(line)
                            extracted_count += 1
                            st.write(f"✅ 抽出（名前）: '{line}'")
                            expect_id = True  # 次はIDを期待
                        else:
                            # 数字の行が来た場合、これは次のIDの可能性
                            st.write(f"🔢 次のID発見: '{line}'")
                            expect_id = False  # 次はクライアント名を期待
            
                st.write(f"📊 このページから {extracted_count} 件のクライアント名を抽出しました")

    except Exception as e:
        st.error(f"クライアント名抽出中にエラーが発生しました: {e}")
        st.exception(e)

    # 結果の表示
    st.write(f"🎯 最終結果: 総 {len(client_names)} 件のクライアント名を抽出")
    if client_names:
        st.write("抽出されたクライアント名:")
        for i, name in enumerate(client_names[:10]):  # 最初の10件を表示
            st.write(f"  {i+1}: {name}")
        if len(client_names) > 10:
            st.write(f"  ... 他 {len(client_names) - 10} 件")

    return client_names

# PDF → Excel 変換 ページ
if page_selection == "PDF → Excel 変換":
    st.markdown('<div class="title">【数出表】PDF → Excelへの変換</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">PDFの数出表をExcelに変換し、同時に盛り付け札を作成します。</div>', unsafe_allow_html=True)

    # UI：PDFファイルアップロード
    uploaded_pdf = st.file_uploader("処理するPDFファイルをアップロードしてください", type="pdf",
                                    help="ここにPDFファイルをドラッグ＆ドロップするか、クリックして選択してください。")

    # ファイル処理とダウンロード表示用のコンテナ
    file_container = st.container()
    download_container = st.container()

    # PDFがアップロードされたら処理を実行
    if uploaded_pdf is not None and st.session_state.template_wb is not None:
        # 処理中の表示
        with file_container:
            file_ext = uploaded_pdf.name.split('.')[-1].lower()
            file_icon = "PDF"
            file_size = len(uploaded_pdf.getvalue()) / 1024

            progress_placeholder = st.empty()
            progress_placeholder.markdown(f"""
            <div class="file-card">
                <div class="file-info">
                    <div class="file-icon">{file_icon}</div>
                    <div class="file-details">
                        <div class="file-name">{uploaded_pdf.name}</div>
                        <div class="file-meta">{file_size:.0f} KB</div>
                    </div>
                </div>
                <div class="loading-spinner"></div>
            </div>
            <div class="progress-bar"><div class="progress-value"></div></div>
            """, unsafe_allow_html=True)

        # PDFのバイナリデータをio.BytesIOに変換
        pdf_bytes_io = io.BytesIO(uploaded_pdf.getvalue())

        # DataFrameへの変換（貼り付け用シート向け）
        df_paste_sheet = None
        with st.spinner("「貼り付け用」データを抽出中..."):
            pdf_bytes_io.seek(0) 
            df_paste_sheet = pdf_to_excel_data_for_paste_sheet(pdf_bytes_io)

        # DataFrameへの変換（注文弁当の抽出シート向け）
        df_bento_sheet = None
        if df_paste_sheet is not None:
            with st.spinner("「注文弁当の抽出」データを抽出中..."):
                try:
                    pdf_bytes_io.seek(0)
                    tables = extract_table_from_pdf_for_bento(pdf_bytes_io)
                    if not tables:
                        st.warning("PDFから表を抽出できませんでした。（注文弁当の抽出）")
                    else:
                        main_table = max(tables, key=lambda t: len(t) * len(t[0])) if tables else []
                        if not main_table:
                            st.warning("メインとなる表が見つかりませんでした。（注文弁当の抽出）")
                        else:
                            anchor_col = find_correct_anchor_for_bento(main_table)
                            if anchor_col == -1:
                                st.warning("「赤」行下の「飯なし」を見つけられませんでした。（注文弁当の抽出）")
                            else:
                                bento_list = extract_bento_range_for_bento(main_table, anchor_col)
                                if not bento_list:
                                    st.warning("弁当範囲を抽出できませんでした。（注文弁当の抽出）")
                                else:
                                    # セッションステートのマスタデータを使用
                                    matched_list = match_bento_names(bento_list, st.session_state.master_df)
                                    output_data_bento = []
                                    for item in matched_list:
                                        match_found = False
                                        match = re.search(r' \(入数: (.+?)\)$', item)
                                        if match:
                                            bento_name = item[:match.start()]
                                            bento_count = match.group(1)
                                            output_data_bento.append([bento_name.strip(), bento_count.strip()])
                                            match_found = True
                                        elif "(未マッチ)" in item:
                                            bento_name = item.replace(" (未マッチ)", "").strip()
                                            bento_count = ""
                                            output_data_bento.append([bento_name, bento_count])
                                            match_found = True
                                        if not match_found:
                                            output_data_bento.append([item.strip(), ""])
                                    df_bento_sheet = pd.DataFrame(output_data_bento, columns=['商品予定名', 'パン箱入数'])
                except Exception as e:
                    st.error(f"「注文弁当の抽出」データ処理中にエラーが発生しました: {e}")
                    st.exception(e)

        # DataFrameへの変換（クライアント抽出シート向け）
        df_client_sheet = None
        if df_paste_sheet is not None:
            with st.spinner("「クライアント抽出」データを抽出中..."):
                try:
                    pdf_bytes_io.seek(0)
                    client_names = extract_client_names_from_pdf(pdf_bytes_io)
                    if not client_names:
                        st.warning("PDFからクライアント名を抽出できませんでした。")
                    else:
                        # クライアント名をDataFrameに変換
                        output_data_client = [[name] for name in client_names]
                        df_client_sheet = pd.DataFrame(output_data_client, columns=['クライアント名'])
                except Exception as e:
                    st.error(f"「クライアント抽出」データ処理中にエラーが発生しました: {e}")
                    st.exception(e)

        # Excelに書き込み
        if df_paste_sheet is not None:
            try:
                with st.spinner("Excelテンプレートにデータを書き込み中..."):
                    try:
                        ws_paste = st.session_state.template_wb["貼り付け用"]
                        # 既存データをクリアしてから書き込む場合は以下のコメントを解除
                        # ws_paste.delete_rows(1, ws_paste.max_row)
                        for r_idx, row in df_paste_sheet.iterrows():
                            for c_idx, value in enumerate(row):
                                ws_paste.cell(row=r_idx + 1, column=c_idx + 1, value=value)
                    except KeyError:
                        st.error("エラー: テンプレートファイルに「貼り付け用」という名前のシートが見つかりません。")
                        st.stop()
                    
                    if df_bento_sheet is not None and not df_bento_sheet.empty:
                        try:
                            ws_bento = st.session_state.template_wb["注文弁当の抽出"]
                            # 既存データをクリアしてから書き込む場合は以下のコメントを解除
                            # ws_bento.delete_rows(1, ws_bento.max_row)
                            for r_idx, row in df_bento_sheet.iterrows():
                                for c_idx, value in enumerate(row):
                                    ws_bento.cell(row=r_idx + 1, column=c_idx + 1, value=value)
                        except KeyError:
                            st.error("エラー: テンプレートファイルに「注文弁当の抽出」という名前のシートが見つかりません。")
                            st.stop()
                    elif df_bento_sheet is not None and df_bento_sheet.empty:
                        st.warning("「注文弁当の抽出」シートに書き込むデータがありませんでした。")
                    else:
                        st.warning("「注文弁当の抽出」データの準備ができませんでした。このシートへの書き込みはスキップされます。")

                    # クライアント抽出シートへの書き込み
                    if df_client_sheet is not None and not df_client_sheet.empty:
                        try:
                            ws_client = st.session_state.template_wb["クライアント抽出"]
                            # 既存データをクリアしてから書き込む場合は以下のコメントを解除
                            # ws_client.delete_rows(1, ws_client.max_row)
                            for r_idx, row in df_client_sheet.iterrows():
                                for c_idx, value in enumerate(row):
                                    ws_client.cell(row=r_idx + 1, column=c_idx + 1, value=value)
                        except KeyError:
                            st.error("エラー: テンプレートファイルに「クライアント抽出」という名前のシートが見つかりません。")
                            st.stop()
                    elif df_client_sheet is not None and df_client_sheet.empty:
                        st.warning("「クライアント抽出」シートに書き込むデータがありませんでした。")
                    else:
                        st.warning("「クライアント抽出」データの準備ができませんでした。このシートへの書き込みはスキップされます。")

                # メモリ上でExcelファイルを生成
                with st.spinner("Excelファイルを生成中..."):
                    output = io.BytesIO()
                    st.session_state.template_wb.save(output)
                    output.seek(0)
                    final_excel_bytes = output.read()

                # 処理完了表示
                with file_container:
                        progress_placeholder.markdown(f"""
                        <div class="file-card">
                            <div class="file-info">
                                <div class="file-icon">{file_icon}</div>
                                <div class="file-details">
                                    <div class="file-name">{uploaded_pdf.name}</div>
                                    <div class="file-meta">{file_size:.0f} KB - 処理完了</div>
                                </div>
                            </div>
                            <div class="check-icon">✓</div>
                        </div>
                        """, unsafe_allow_html=True)

                # ダウンロードリンクの生成
                with download_container:
                    st.markdown('<div class="separator"></div>', unsafe_allow_html=True)

                    original_pdf_name = os.path.splitext(uploaded_pdf.name)[0]
                    output_filename = f"{original_pdf_name}_Processed.xlsm"
                    excel_size = len(final_excel_bytes) / 1024
                    b64 = base64.b64encode(final_excel_bytes).decode('utf-8')

                    mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12"

                    href = f"""
                    <a href="data:{mime_type};base64,{b64}" download="{output_filename}" class="download-card">
                        <div class="download-info">
                            <div class="download-icon">XLSM</div>
                            <div class="download-details">
                                <div class="download-name">{output_filename}</div>
                                <div class="download-meta">Excel (マクロ有効)・{excel_size:.0f} KB</div>
                            </div>
                        </div>
                        <div class="download-button-imitation">
                            <span class="download-button-icon">↓</span>
                            Download
                        </div>
                    </a>
                    """
                    st.markdown(href, unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Excelファイルへの書き込みまたは生成中にエラーが発生しました: {e}")
                st.exception(e)
                with file_container:
                        progress_placeholder.markdown(f"エラー発生: {e}", unsafe_allow_html=True)

# クライアント情報抽出 ページ
elif page_selection == "クライアント情報抽出":
    st.markdown('<div class="title">クライアント情報抽出</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">PDFからクライアント名と給食の数を抽出してCSVファイルに出力します。</div>', unsafe_allow_html=True)

    # PDFファイルアップロード
    uploaded_file = st.file_uploader("PDFファイルをアップロードしてください", type="pdf")
    
    if uploaded_file is not None:
        # クライアント情報を抽出
        client_data = extract_client_and_meal_info_from_pdf(uploaded_file)
        
        if client_data:
            # データをDataFrameに変換
            df = export_client_data_to_csv(client_data)
            
            # データの表示
            st.success(f"データを抽出しました: {len(client_data)} 件")
            st.dataframe(df)
            
            # CSVデータを生成
            csv_data = df.to_csv(index=False, encoding='utf-8-sig')
            
            # ダウンロードボタン
            st.download_button(
                label="📥 CSVファイルをダウンロード",
                data=csv_data,
                file_name="client_meal_data.csv",
                mime="text/csv",
                help="クライアント情報と給食の数をCSVファイルとしてダウンロードします"
            )
        else:
            st.warning("クライアント情報を抽出できませんでした。PDFの形式を確認してください。")

# マスタ設定 ページ
elif page_selection == "マスタ設定":
    st.markdown('<div class="title">マスタデータ設定</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">商品マスタのCSVファイルをアップロードして更新します。現在のマスタデータも確認できます。</div>', unsafe_allow_html=True)

    # --- マスタCSVのファイルパス ---
    master_csv_path = "商品マスタ一覧.csv"

    # --- UI: 新しいマスタCSVのアップロード ---
    st.markdown("#### 新しいマスタをアップロード")
    uploaded_master_csv = st.file_uploader(
        "商品マスタ一覧.csv をアップロードしてください",
        type="csv",
        help="ヘッダーには '商品予定名' と 'パン箱入数' を含めてください。"
    )

    if uploaded_master_csv is not None:
        try:
            # --- アップロードされたCSVをDataFrameとして読み込む ---
            new_master_df = None
            # BOM付きUTF-8、Shift_JISなど、複数のエンコーディングを試す
            encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
            for encoding in encodings:
                try:
                    uploaded_master_csv.seek(0) # ファイルポインタを先頭に戻す
                    temp_df = pd.read_csv(uploaded_master_csv, encoding=encoding)
                    # 必須カラムの存在チェック
                    if '商品予定名' in temp_df.columns and 'パン箱入数' in temp_df.columns:
                        new_master_df = temp_df
                        st.info(f"アップロードされたファイルを {encoding} で読み込みました。")
                        break
                    else:
                        st.warning(f"{encoding} で読み込みましたが、必須列（'商品予定名', 'パン箱入数'）が見つかりません。")

                except (UnicodeDecodeError, pd.errors.ParserError):
                    continue # 次のエンコーディングを試す
                except Exception as e:
                    st.error(f"ファイルの読み込み中に予期せぬエラーが発生しました: {e}")
                    break

            if new_master_df is not None:
                # --- セッションステートを更新 ---
                st.session_state.master_df = new_master_df

                # --- ✅ CSVファイルに上書き保存 ---
                try:
                    # UTF-8 (BOM付き)で保存。Excelでの文字化けを防ぐ
                    new_master_df.to_csv(master_csv_path, index=False, encoding='utf-8-sig')
                    st.success(f"✅ マスタデータを更新し、'{master_csv_path}' に上書き保存しました。")
                    st.info("アプリを再起動しても、このマスタが読み込まれます。")

                except Exception as e:
                    st.error(f"マスタファイルの保存中にエラーが発生しました: {e}")
                    st.exception(e)

            else:
                st.error("アップロードされたCSVファイルを正しく読み込めませんでした。ファイルの形式（必須列の有無）やエンコーディングを確認してください。")

        except Exception as e:
            st.error(f"マスタ更新処理中にエラーが発生しました: {e}")
            st.exception(e)

    st.markdown('<div class="separator"></div>', unsafe_allow_html=True)

    # --- 現在のマスタデータを表示 ---
    st.markdown("#### 現在のマスタデータ")
    if 'master_df' in st.session_state and not st.session_state.master_df.empty:
        st.dataframe(st.session_state.master_df, use_container_width=True)
    else:
        st.warning("現在、マスタデータが読み込まれていません。")

# メインコンテナ終了
st.markdown('</div>', unsafe_allow_html=True)
