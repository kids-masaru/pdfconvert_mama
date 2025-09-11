# pdf_utils.py - 完全版（既存の関数 + 新しい改善された関数）

import pdfplumber
import pandas as pd
import re
from typing import List, Tuple, Any, Optional

# ===============================================
# 既存の関数（元のpdf_utils.pyから必要な関数を追加）
# ===============================================

def safe_write_df(worksheet, df, start_row=1):
    """DataFrameを安全にワークシートに書き込む"""
    try:
        for r_idx, row in df.iterrows():
            for c_idx, value in enumerate(row):
                worksheet.cell(row=r_idx + start_row, column=c_idx + 1, value=value)
    except Exception as e:
        print(f"DataFrame書き込みエラー: {e}")

def pdf_to_excel_data_for_paste_sheet(pdf_bytes_io):
    """従来のPDFからExcelへのデータ変換（フォールバック用）"""
    try:
        with pdfplumber.open(pdf_bytes_io) as pdf:
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
            
            if not all_text.strip():
                return pd.DataFrame()
            
            # 行ごとに分割
            lines = all_text.split('\n')
            data_rows = []
            
            for line in lines:
                if line.strip():
                    # スペースで分割（複数スペースを考慮）
                    cells = re.split(r'\s{2,}', line.strip())
                    if len(cells) > 1:  # 複数列がある行のみ
                        data_rows.append(cells)
            
            if data_rows:
                # 最大列数に合わせる
                max_cols = max(len(row) for row in data_rows)
                normalized_data = []
                for row in data_rows:
                    while len(row) < max_cols:
                        row.append("")
                    normalized_data.append(row)
                
                return pd.DataFrame(normalized_data)
            
            return pd.DataFrame()
            
    except Exception as e:
        print(f"従来のPDF変換エラー: {e}")
        return pd.DataFrame()

def extract_table_from_pdf_for_bento(pdf_bytes_io):
    """弁当用テーブル抽出"""
    try:
        with pdfplumber.open(pdf_bytes_io) as pdf:
            tables = []
            for page in pdf.pages:
                page_tables = page.extract_tables()
                if page_tables:
                    tables.extend(page_tables)
            return tables
    except Exception as e:
        print(f"弁当テーブル抽出エラー: {e}")
        return []

def find_correct_anchor_for_bento(table):
    """弁当用アンカー検索"""
    try:
        if not table:
            return -1
        
        # テーブルの最初の行からアンカー列を探す
        header_row = table[0] if table else []
        for i, cell in enumerate(header_row):
            if cell and ('商品' in str(cell) or '弁当' in str(cell) or '品名' in str(cell)):
                return i
        
        return 0  # デフォルトは最初の列
    except Exception:
        return -1

def extract_bento_range_for_bento(table, anchor_col):
    """弁当範囲抽出"""
    try:
        bento_list = []
        if not table or anchor_col < 0:
            return bento_list
        
        for row in table[1:]:  # ヘッダー行をスキップ
            if len(row) > anchor_col and row[anchor_col]:
                bento_name = str(row[anchor_col]).strip()
                if bento_name:
                    bento_list.append(bento_name)
        
        return bento_list
    except Exception as e:
        print(f"弁当範囲抽出エラー: {e}")
        return []

def match_bento_names(bento_list, master_df):
    """弁当名マッチング"""
    try:
        matched_list = []
        
        if master_df.empty or '商品予定名' not in master_df.columns:
            return [f"{name} (未マッチ)" for name in bento_list]
        
        master_names = master_df['商品予定名'].tolist()
        
        for bento_name in bento_list:
            matched = False
            for master_name in master_names:
                if str(master_name).strip() == str(bento_name).strip():
                    # 入数情報を追加
                    if 'パン箱入数' in master_df.columns:
                        matched_row = master_df[master_df['商品予定名'] == master_name]
                        if not matched_row.empty:
                            iri_value = matched_row.iloc[0]['パン箱入数']
                            matched_list.append(f"{bento_name} (入数: {iri_value})")
                            matched = True
                            break
            
            if not matched:
                matched_list.append(f"{bento_name} (未マッチ)")
        
        return matched_list
    except Exception as e:
        print(f"弁当名マッチングエラー: {e}")
        return [f"{name} (エラー)" for name in bento_list]

def extract_detailed_client_info_from_pdf(pdf_bytes_io):
    """詳細クライアント情報抽出"""
    try:
        with pdfplumber.open(pdf_bytes_io) as pdf:
            client_data = []
            
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        # クライアント情報のパターンを検索
                        if any(keyword in line for keyword in ['株式会社', '有限会社', '会社', '様']):
                            client_data.append({
                                'client_info': line.strip(),
                                'page': len(client_data) + 1
                            })
            
            return client_data
    except Exception as e:
        print(f"クライアント情報抽出エラー: {e}")
        return []

def export_detailed_client_data_to_dataframe(client_data):
    """クライアントデータをDataFrameに変換"""
    try:
        if not client_data:
            return pd.DataFrame()
        
        df_data = []
        for item in client_data:
            df_data.append([
                item.get('client_info', ''),
                item.get('page', '')
            ])
        
        return pd.DataFrame(df_data, columns=['クライアント情報', 'ページ'])
    except Exception as e:
        print(f"クライアントDataFrame変換エラー: {e}")
        return pd.DataFrame()

# ===============================================
# 新しい改善された関数
# ===============================================

def extract_text_with_numbers(pdf_bytes_io) -> str:
    """PDFから文字と数字の両方を確実に抽出する"""
    try:
        with pdfplumber.open(pdf_bytes_io) as pdf:
            full_text = ""
            for page in pdf.pages:
                # 複数の抽出方法を試行
                text_methods = [
                    # 方法1: 通常のテキスト抽出
                    lambda p: p.extract_text() or "",
                    # 方法2: 文字レベルでの抽出
                    lambda p: extract_chars_with_positions(p),
                    # 方法3: テーブル形式での抽出
                    lambda p: extract_from_table_structure(p)
                ]
                
                page_text = ""
                for method in text_methods:
                    try:
                        extracted = method(page)
                        if extracted and len(extracted) > len(page_text):
                            page_text = extracted
                    except Exception:
                        continue
                
                full_text += page_text + "\n"
            
            return full_text
    except Exception as e:
        print(f"PDF抽出エラー: {e}")
        return ""

def extract_chars_with_positions(page) -> str:
    """文字レベルでの詳細な抽出"""
    try:
        chars = page.chars
        if not chars:
            return ""
        
        # 位置情報を使って文字を並べ替え
        sorted_chars = sorted(chars, key=lambda x: (x['y0'], x['x0']))
        
        text_lines = []
        current_line = []
        current_y = None
        
        for char in sorted_chars:
            # 新しい行の判定
            if current_y is None or abs(char['y0'] - current_y) > 5:  # 5ポイント以上の差で新しい行
                if current_line:
                    text_lines.append(''.join(current_line))
                current_line = [char['text']]
                current_y = char['y0']
            else:
                current_line.append(char['text'])
        
        if current_line:
            text_lines.append(''.join(current_line))
        
        return '\n'.join(text_lines)
    except Exception:
        return ""

def extract_from_table_structure(page) -> str:
    """テーブル構造からの抽出"""
    try:
        tables = page.extract_tables()
        if not tables:
            return ""
        
        text_lines = []
        for table in tables:
            for row in table:
                if row:
                    # None値を空文字に変換し、数値も文字列として保持
                    clean_row = [str(cell) if cell is not None else "" for cell in row]
                    text_lines.append('\t'.join(clean_row))
        
        return '\n'.join(text_lines)
    except Exception:
        return ""

def improved_pdf_to_excel_data_for_paste_sheet(pdf_bytes_io) -> pd.DataFrame:
    """改善されたPDFからExcelへのデータ変換"""
    try:
        # テキスト抽出
        text_content = extract_text_with_numbers(pdf_bytes_io)
        
        if not text_content.strip():
            return pd.DataFrame()
        
        # 行ごとに分割
        lines = text_content.split('\n')
        
        # データを格納するリスト
        data_rows = []
        
        for line in lines:
            if line.strip():
                # タブ区切りの場合
                if '\t' in line:
                    cells = line.split('\t')
                else:
                    # スペース区切りの場合（複数スペースを考慮）
                    cells = re.split(r'\s{2,}', line.strip())
                
                # 数値の正規化処理
                processed_cells = []
                for cell in cells:
                    cell = cell.strip()
                    if cell:
                        # 数値パターンのチェック
                        if re.match(r'^[\d,.-]+$', cell):
                            # カンマを除去して数値として処理
                            try:
                                if '.' in cell:
                                    processed_cells.append(float(cell.replace(',', '')))
                                else:
                                    processed_cells.append(int(cell.replace(',', '')))
                            except ValueError:
                                processed_cells.append(cell)
                        else:
                            processed_cells.append(cell)
                
                if processed_cells:
                    data_rows.append(processed_cells)
        
        if not data_rows:
            return pd.DataFrame()
        
        # 最大列数を取得
        max_cols = max(len(row) for row in data_rows)
        
        # 不足している列を空文字で埋める
        normalized_rows = []
        for row in data_rows:
            while len(row) < max_cols:
                row.append("")
            normalized_rows.append(row)
        
        return pd.DataFrame(normalized_rows)
        
    except Exception as e:
        print(f"PDF変換エラー: {e}")
        return pd.DataFrame()

def extract_numbers_specifically(text: str) -> List[str]:
    """テキストから数値を特別に抽出"""
    # 数値パターン（整数、小数、カンマ区切り）
    number_patterns = [
        r'\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b',  # 1,000.5 形式
        r'\b\d+\.\d+\b',                       # 123.45 形式
        r'\b\d+\b'                            # 123 形式
    ]
    
    numbers = []
    for pattern in number_patterns:
        matches = re.findall(pattern, text)
        numbers.extend(matches)
    
    return numbers

def debug_pdf_content(pdf_bytes_io) -> dict:
    """PDFの内容をデバッグ用に詳細表示"""
    debug_info = {
        'pages': 0,
        'total_chars': 0,
        'numbers_found': [],
        'text_sample': ""
    }
    
    try:
        with pdfplumber.open(pdf_bytes_io) as pdf:
            debug_info['pages'] = len(pdf.pages)
            
            for i, page in enumerate(pdf.pages):
                # 文字レベルの情報
                chars = page.chars
                debug_info['total_chars'] += len(chars) if chars else 0
                
                # テキスト抽出のサンプル
                text = page.extract_text() or ""
                if i == 0:  # 最初のページのサンプル
                    debug_info['text_sample'] = text[:500]
                
                # 数値の検出
                numbers = extract_numbers_specifically(text)
                debug_info['numbers_found'].extend(numbers)
    
    except Exception as e:
        debug_info['error'] = str(e)
    
    return debug_info
