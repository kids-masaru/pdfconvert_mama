# pdf_utils.py

import pandas as pd
import pdfplumber
import re
import unicodedata
from typing import List, Dict, Any

def safe_write_df(worksheet, df, start_row=1):
    """DataFrameをExcelシートに安全に書き込む"""
    num_cols = df.shape[1]
    if worksheet.max_row >= start_row:
        for row_idx in range(start_row, worksheet.max_row + 2):
            for col_idx in range(1, num_cols + 2):
                worksheet.cell(row=row_idx, column=col_idx).value = None
    for r_idx, row_data in enumerate(df.itertuples(index=False), start=start_row):
        for c_idx, value in enumerate(row_data, start=1):
            worksheet.cell(row=r_idx, column=c_idx, value=value)

def match_bento_data(pdf_bento_list: List[str], master_df: pd.DataFrame) -> List[List[str]]:
    """
    PDFの弁当名リストを商品マスタと照合し、関連データを返す。
    CSVのヘッダー問題をここで吸収し、安全な列名でデータを取得する。
    """
    if master_df is None or master_df.empty:
        return [[name, "", "", ""] for name in pdf_bento_list]

    # --- ★最重要：CSVの全ヘッダー名から見えないスペースを除去 ---
    master_df.columns = master_df.columns.str.strip()

    required_cols = ['商品予定名', 'パン箱入数', 'クラス分け名称4', 'クラス分け名称5']
    if not all(col in master_df.columns for col in required_cols):
        missing = ", ".join([col for col in required_cols if col not in master_df.columns])
        return [[name, "", f"マスタ列不足: {missing}", ""] for name in pdf_bento_list]
    
    # 必要な列のデータを安全に抽出
    master_tuples = master_df[required_cols].astype(str).to_records(index=False).tolist()
    matched_results = []
    
    norm_master = [
        (unicodedata.normalize('NFKC', name).replace(" ", ""), name, pan_box, c4, c5)
        for name, pan_box, c4, c5 in master_tuples
    ]

    for pdf_name in pdf_bento_list:
        pdf_name_stripped = pdf_name.strip()
        norm_pdf = unicodedata.normalize('NFKC', pdf_name_stripped).replace(" ", "")
        result_data = [pdf_name_stripped, "", "", ""]
        best_match = None
        
        # 1. 完全一致で検索
        for norm_m, orig_m, pan_box, c4, c5 in norm_master:
            if norm_m == norm_pdf:
                best_match = [orig_m, pan_box, c4, c5]
                break
        
        # 2. 部分一致で検索
        if not best_match:
            candidates = []
            for norm_m, orig_m, pan_box, c4, c5 in norm_master:
                if norm_m and norm_m in norm_pdf:
                    candidates.append((orig_m, pan_box, c4, c5))
            if candidates:
                best_match = max(candidates, key=lambda x: len(x[0]))

        if best_match:
            result_data = best_match
        
        matched_results.append(result_data)
        
    return matched_results

# ──────────────────────────────────────────────
# 以下の関数は変更ありません
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
                    if '10001' in ''.join(str(c) for c in row if c): break
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
    except Exception:
        pass
    return client_data

def extract_meal_numbers_from_row(rows, row_idx, client_id, client_name):
    client_info = {'client_id': client_id, 'client_name': client_name, 'student_meals': [], 'teacher_meals': []}
    rows_to_check = []
    for i in range(max(0, row_idx - 3), min(len(rows), row_idx + 3)):
        if i < len(rows) and rows[i]:
            left_cell = str(rows[i][0]).strip()
            if left_cell == client_id: rows_to_check.append(('id', rows[i]))
            elif left_cell == client_name: rows_to_check.append(('name', rows[i]))
    all_numbers = []
    for row_type, row in rows_to_check:
        for cell in row[1:]:
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

def extract_text_with_layout(page) -> List[List[str]]:
    words = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False)
    if not words: return []
    boundaries = get_vertical_boundaries(page)
    if len(boundaries) < 2:
        text = page.extract_text(layout=False, x_tolerance=3, y_tolerance=3)
        return [[line] for line in text.split('\n') if line.strip()] if text else []
    row_groups = get_line_groups(words, y_tolerance=1.5)
    result_rows = []
    for group in row_groups:
        sorted_group = sorted(group, key=lambda w: w['x0'])
        columns = split_line_using_boundaries(sorted_group, boundaries)
        if any(cell.strip() for cell in columns):
            result_rows.append(columns)
    return result_rows

def get_line_groups(words: List[Dict[str, Any]], y_tolerance: float = 1.2) -> List[List[Dict[str, Any]]]:
    if not words: return []
    sorted_words = sorted(words, key=lambda w: w['top'])
    groups, current_group = [], [sorted_words[0]]
    for word in sorted_words[1:]:
        if abs(word['top'] - current_group[-1]['top']) <= y_tolerance:
            current_group.append(word)
        else:
            groups.append(sorted(current_group, key=lambda w: w['x0']))
            current_group = [word]
    groups.append(sorted(current_group, key=lambda w: w['x0']))
    return groups

def get_vertical_boundaries(page, tolerance: float = 2) -> List[float]:
    lines = page.lines
    v_lines_x = sorted(list(set(round(line['x0'], 1) for line in lines if line['height'] > 0 and line['width'] < tolerance)))
    words = page.extract_words()
    if not words: return v_lines_x
    doc_left = min(word['x0'] for word in words)
    doc_right = max(word['x1'] for word in words)
    boundaries = sorted(list(set([round(doc_left, 1)] + v_lines_x + [round(doc_right, 1)])))
    merged = []
    if boundaries:
        merged.append(boundaries[0])
        for b in boundaries[1:]:
            if b - merged[-1] > tolerance * 2:
                merged.append(b)
    return merged

def split_line_using_boundaries(line: List[Dict[str, Any]], boundaries: List[float]) -> List[str]:
    columns = [""] * (len(boundaries) - 1)
    for word in line:
        word_center = (word['x0'] + word['x1']) / 2
        for i in range(len(boundaries) - 1):
            if boundaries[i] <= word_center < boundaries[i+1]:
                columns[i] = (columns[i] + " " + word["text"]).strip()
                break
    return columns

def pdf_to_excel_data_for_paste_sheet(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            if not pdf.pages: return None
            page = pdf.pages[0]
            rows = extract_text_with_layout(page)
            if not rows: return None
            df = pd.DataFrame(rows)
            df.replace({None: ""}, inplace=True)
            return df
    except Exception:
        return None

def extract_table_from_pdf_for_bento(pdf_file_obj):
    tables = []
    with pdfplumber.open(pdf_file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text or not any(kw in text for kw in ["園名", "飯あり", "キャラ弁"]): continue
            if not page.lines: continue
            table = page.extract_table({"vertical_strategy": "lines", "horizontal_strategy": "lines"})
            if table: tables.append(table)
    return tables

def find_correct_anchor_for_bento(table, target_row_text="赤"):
    for r_idx, row in enumerate(table):
        if target_row_text in ''.join(str(c) for c in row if c):
            if r_idx + 1 < len(table):
                for c_idx, cell in enumerate(table[r_idx + 1]):
                    if cell and "飯なし" in cell: return c_idx
    return -1

def extract_bento_range_for_bento(table, start_col):
    bento_list, end_col = [], -1
    for row in table:
        if "おやつ" in ''.join(str(c) for c in row if c):
            for c_idx, cell in enumerate(row):
                if cell and "おやつ" in cell: end_col = c_idx; break
            if end_col != -1: break
    if end_col == -1 or start_col >= end_col: return []
    header_row_idx = -1
    for r_idx, row in enumerate(table):
        if any(c and "飯なし" in c for c in row):
            if r_idx > 0: header_row_idx = r_idx - 1; break
    if header_row_idx == -1: return []
    header_row = table[header_row_idx]
    for col in range(start_col + 1, end_col):
        cell_text = header_row[col] if col < len(header_row) else ""
        if cell_text and str(cell_text).strip(): bento_list.append(str(cell_text).strip())
    return bento_list
