import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import base64
import os
from typing import List, Dict, Any
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet # Worksheetのインポートを追加
# from copy import copy # 書式コピー用のimport (現在はコメントアウト)

# ----------------------------
# ページ設定（アイコン指定：ブラウザタブ・ブックマーク用）
# ----------------------------
# アイコンファイルが存在する場合のみ設定
icon_path = "icon.png"
page_icon_value = icon_path if os.path.exists(icon_path) else None

st.set_page_config(
    page_title="【数出表】PDF → Excelへの変換",
    layout="centered",
    page_icon=page_icon_value
)

# ----------------------------
# UIのスタイル設定（洗練されたモダンデザイン - 暖色系背景）
# ----------------------------
# (スタイル設定は変更なしのため省略)
st.markdown("""
    <style>
        /* Google FontsのInter, Robotoをインポート */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Roboto:wght@300;400;500&display=swap');

        /* アプリ全体の背景とフォント - 薄いオレンジ系の背景 */
        .stApp {
            background: #fff5e6;
            font-family: 'Inter', sans-serif;
        }

        /* タイトル */
        .title {
            font-size: 1.5rem;
            font-weight: 600;
            color: #333;
            margin-bottom: 5px;
        }

        /* サブタイトル */
        .subtitle {
            font-size: 0.9rem;
            color: #666;
            margin-bottom: 25px;
        }

        /* ファイルアップローダーのカスタマイズ */
        [data-testid="stFileUploader"] {
            background: #ffffff;
            border-radius: 10px;
            border: 1px dashed #d0d0d0;
            padding: 30px 20px;
            margin: 20px 0;
        }

        [data-testid="stFileUploader"] label {
            display: none;
        }

        [data-testid="stFileUploader"] section {
            border: none !important;
            background: transparent !important;
        }

        /* ファイル情報カード */
        .file-card {
            background: white;
            border-radius: 8px;
            padding: 12px 16px;
            margin: 15px 0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            display: flex;
            align-items: center;
            justify-content: space-between;
            border: 1px solid #eaeaea;
        }

        .file-info {
            display: flex;
            align-items: center;
        }

        .file-icon {
            width: 36px;
            height: 36px;
            border-radius: 6px;
            background-color: #f44336; /* PDF icon color */
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 12px;
            color: white;
            font-weight: 500;
            font-size: 14px;
        }
        .excel-icon { /* Added for Excel download card */
             width: 40px;
             height: 40px;
             border-radius: 8px;
             background-color: #ff9933; /* Excel icon color */
             display: flex;
             align-items: center;
             justify-content: center;
             margin-right: 12px;
             color: white;
             font-weight: 500;
             font-size: 16px;
        }


        .file-details {
            display: flex;
            flex-direction: column;
        }

        .file-name {
            font-weight: 500;
            color: #333;
            font-size: 0.9rem;
            margin-bottom: 3px;
        }

        .file-meta {
            font-size: 0.75rem;
            color: #888;
        }

        /* ローディングアニメーション */
        .loading-spinner {
            width: 20px;
            height: 20px;
            border: 2px solid rgba(0,0,0,0.1);
            border-radius: 50%;
            border-top-color: #ff9933; /* Spinner color */
            animation: spin 1s linear infinite;
        }

        .check-icon {
            color: #ff9933; /* Checkmark color */
            font-size: 20px;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        /* 進行状況バー (Optional, kept for visual feedback during processing) */
        .progress-bar {
            height: 4px;
            background-color: #e0e0e0;
            border-radius: 2px;
            width: 100%;
            margin-top: 10px;
            overflow: hidden; /* Ensure progress value stays within bounds */
        }

        .progress-value {
            height: 100%;
            background-color: #ff9933; /* Progress bar color */
            border-radius: 2px;
            width: 0%; /* Start at 0% */
            transition: width 0.5s ease-in-out; /* Smooth transition */
        }
        .progress-value.done {
             width: 100%; /* Set to 100% when done */
        }


        /* ダウンロードカード */
        .download-card {
            background: white;
            border-radius: 8px;
            padding: 16px;
            margin: 20px 0;
            box-shadow: 0 2px 5px rgba(0,0,0,0.08);
            display: flex;
            align-items: center;
            justify-content: space-between;
            border: 1px solid #eaeaea;
            transition: all 0.2s ease;
            cursor: pointer;
            text-decoration: none; /* Remove underline from link */
            color: inherit; /* Inherit text color */
        }

        .download-card:hover {
            box-shadow: 0 4px 12px rgba(0,0,0,0.12);
            background-color: #fffaf0; /* Light orange on hover */
            transform: translateY(-2px);
        }

        .download-info {
            display: flex;
            align-items: center;
        }

        /* download-icon is now excel-icon */

        .download-details {
            display: flex;
            flex-direction: column;
        }

        .download-name {
            font-weight: 500;
            color: #333;
            font-size: 0.9rem;
            margin-bottom: 3px;
        }

        .download-meta {
            font-size: 0.75rem;
            color: #888;
        }

        .download-button {
            background-color: #ff9933; /* Button color */
            color: white;
            border: none;
            border-radius: 6px;
            padding: 8px 16px;
            font-size: 0.85rem;
            font-weight: 500;
            cursor: pointer;
            transition: background-color 0.2s;
            display: flex;
            align-items: center;
        }

        .download-button:hover {
            background-color: #e68a00; /* Darker orange on hover */
        }

        .download-button-icon {
            margin-right: 6px;
        }

        /* Hide default Streamlit spinner text */
        .stSpinner > div {
            /* display: none; */ /* Keep spinner text for clarity */
        }

        /* Adjust padding for Streamlit elements if needed */
        .css-1544g2n { /* Might change with Streamlit versions */
            padding-top: 2rem;
        }

        .css-18e3th9 { /* Might change with Streamlit versions */
            padding-top: 2rem;
        }

        /* セパレーター */
        .separator {
            height: 1px;
            background-color: #ffe0b3; /* Lighter separator color */
            margin: 25px 0;
        }
    </style>
""", unsafe_allow_html=True)

# メインコンテナ開始
st.markdown('<div class="main-container">', unsafe_allow_html=True)

# タイトルとサブタイトル
st.markdown('<div class="title">【数出表】PDF → Excelへの変換＆コピー</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">PDFの数出表をExcelに変換し、指定シートの値を別ファイルにコピーします。</div>', unsafe_allow_html=True)

# ----------------------------
# PDF→Excel変換用の関数群
# ----------------------------
def is_number(text: str) -> bool:
    """文字列が数値かどうかを判定する"""
    return bool(re.match(r'^\d+$', text.strip()))

def get_line_groups(words: List[Dict[str, Any]], y_tolerance: float = 1.2) -> List[List[Dict[str, Any]]]:
    """PDF上の単語を行ごとにグループ化する"""
    if not words:
        return []
    # Sort words primarily by 'top' (y-coordinate), then by 'x0' (x-coordinate) for stability
    sorted_words = sorted(words, key=lambda w: (w['top'], w['x0']))
    groups = []
    if not sorted_words:
        return groups

    current_group = [sorted_words[0]]
    current_top = sorted_words[0]['top']

    for word in sorted_words[1:]:
        # Check if the vertical distance is within tolerance
        if abs(word['top'] - current_top) <= y_tolerance:
            current_group.append(word)
        else:
            # Finalize the current group (sorted horizontally) and start a new one
            groups.append(sorted(current_group, key=lambda w: w['x0']))
            current_group = [word]
            current_top = word['top']

    # Add the last group (sorted horizontally)
    groups.append(sorted(current_group, key=lambda w: w['x0']))
    return groups


def get_vertical_boundaries(page, tolerance: float = 2) -> List[float]:
    """PDFのページ内から、縦線とみなせる線のx座標を抽出する"""
    vertical_lines = []
    # Extract vertical lines based on line objects
    for line in page.lines:
        # Check if it's primarily vertical and has some length
        if abs(line['x0'] - line['x1']) < tolerance and abs(line['top'] - line['bottom']) > 5:
             vertical_lines.append((line['x0'] + line['x1']) / 2)

    # Consider table boundaries as potential lines too
    for table in page.find_tables():
        for x in table.bbox[0::2]: # x0, x1 coordinates of the table bbox
             vertical_lines.append(x)

    # Deduplicate and sort
    vertical_lines = sorted(list(set(round(x, 1) for x in vertical_lines)))

    words = page.extract_words(x_tolerance=1, y_tolerance=1) # Use tighter tolerance for boundary detection
    if not words:
        # If no words, use page boundaries if lines are also absent
        return vertical_lines if vertical_lines else [page.bbox[0], page.bbox[2]]

    # Include boundaries based on word positions if lines are sparse
    left_boundary = min(word['x0'] for word in words)
    right_boundary = max(word['x1'] for word in words)

    # Combine line-based and word-based boundaries
    boundaries = sorted(list(set([left_boundary] + vertical_lines + [right_boundary])))

    # Filter out boundaries that are too close together
    filtered_boundaries = []
    if boundaries:
        filtered_boundaries.append(boundaries[0])
        for i in range(1, len(boundaries)):
            if boundaries[i] - filtered_boundaries[-1] > tolerance * 2: # Increase gap tolerance
                 filtered_boundaries.append(boundaries[i])

    # Ensure page edges are included if not already present
    if not filtered_boundaries or filtered_boundaries[0] > page.bbox[0] + tolerance:
         filtered_boundaries.insert(0, page.bbox[0])
    if not filtered_boundaries or filtered_boundaries[-1] < page.bbox[2] - tolerance:
         filtered_boundaries.append(page.bbox[2])

    return sorted(list(set(round(b,1) for b in filtered_boundaries)))


def split_line_using_boundaries(sorted_words: List[Dict[str, Any]], boundaries: List[float]) -> List[str]:
    """同一行内の単語をセルごとに分割する"""
    columns = [""] * (len(boundaries) - 1)
    for word in sorted_words:
        word_center = (word['x0'] + word['x1']) / 2
        for i in range(len(boundaries) - 1):
            left = boundaries[i]
            right = boundaries[i+1]
            # Assign word to the column where its center falls
            # Or if it significantly overlaps
            overlap_threshold = min(5, (right - left) * 0.1) # Overlap threshold
            if (left <= word_center < right) or \
               (word['x0'] < right - overlap_threshold and word['x1'] > left + overlap_threshold):
                columns[i] += word['text'] + " "
                break # Assign to first matching column

    # Trim whitespace from each column
    return [col.strip() for col in columns]

def extract_text_with_layout(page) -> List[List[str]]:
    """PDFページからセル分割されたテキストデータを抽出する"""
    # Use slightly more generous tolerances for word extraction
    words = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=True)
    if not words:
        # st.warning("ページから単語が抽出されませんでした。") # Changed to debug log
        # print("DEBUG: No words extracted from page.")
        return []

    boundaries = get_vertical_boundaries(page)
    if not boundaries or len(boundaries) < 2:
         # st.warning("表の縦境界線を検出できませんでした。レイアウトが崩れる可能性があります。") # Changed to debug log
         # print("DEBUG: Could not detect vertical boundaries reliably.")
         # Fallback: treat the whole line as one cell
         boundaries = [page.bbox[0], page.bbox[2]]

    # Use slightly more generous tolerance for line grouping
    row_groups = get_line_groups(words, y_tolerance=3)
    result_rows = []

    for group in row_groups:
        # Words within a group are already sorted by x0 in get_line_groups
        columns = split_line_using_boundaries(group, boundaries)
        # Only add rows that contain some non-empty cell
        if any(str(cell).strip() for cell in columns if cell is not None): # Check for None
             result_rows.append(columns)

    return result_rows

def remove_extra_empty_columns(rows: List[List[str]]) -> List[List[str]]:
    """すべての行で空の列を削除する"""
    if not rows:
        return rows

    max_cols = 0
    for row in rows:
        # Ensure row is not None and is iterable
        if row:
             max_cols = max(max_cols, len(row))

    if max_cols == 0:
        return rows

    # Pad rows to have the same number of columns before checking
    padded_rows = []
    for row in rows:
        if row: # Ensure row is not None
            padded_rows.append(row + [''] * (max_cols - len(row)))
        else:
            padded_rows.append([''] * max_cols) # Add empty row if original was None

    keep_indices = []
    for col_idx in range(max_cols):
        # Check if any cell in this column index has content (and is not None)
        if any(str(padded_rows[row_idx][col_idx]).strip() for row_idx in range(len(padded_rows)) if padded_rows[row_idx][col_idx] is not None):
            keep_indices.append(col_idx)

    # Create new rows with only the columns to keep
    new_rows = []
    for row in padded_rows:
         # Ensure row is indexable and create new row safely
         new_row = [row[i] for i in keep_indices if i < len(row)]
         new_rows.append(new_row)


    # Remove trailing empty rows if any were created
    while new_rows and not any(str(cell).strip() for cell in new_rows[-1] if cell is not None):
        new_rows.pop()


    return new_rows


def format_excel_worksheet(worksheet: Worksheet):
     """Excelワークシートの書式設定（列幅・行高さ） - openpyxl用"""
     if not isinstance(worksheet, Worksheet):
         print(f"DEBUG: Invalid worksheet passed to format_excel_worksheet: {type(worksheet)}")
         return # Exit if not a valid worksheet

     for col_cells in worksheet.columns:
         max_length = 0
         # Check if col_cells is not empty and contains cells
         if not col_cells or not hasattr(col_cells[0], 'column_letter'):
             continue # Skip if column is empty or invalid

         column = col_cells[0].column_letter # Get the column name

         for cell in col_cells:
             try: # Necessary to avoid error on empty or invalid cells
                 if cell and cell.value is not None: # Check if cell and value exist
                     value_str = str(cell.value)
                     # Consider line breaks for length calculation
                     cell_len = max(len(line) for line in value_str.split('\n'))
                     if cell_len > max_length:
                         max_length = cell_len
             except Exception as e:
                 # print(f"DEBUG: Error processing cell {cell.coordinate} for width: {e}")
                 pass # Ignore errors for individual cells

         # Set a minimum width and calculate adjusted width
         adjusted_width = max(10, (max_length + 2) * 1.2) # Min width 10
         worksheet.column_dimensions[column].width = min(adjusted_width, 60) # Max width 60

     for row_dim in worksheet.row_dimensions.values():
         # Reset height first or use a default
         row_dim.height = 15 # Default height

     # Iterate again to set height based on content
     for row in worksheet.iter_rows():
          max_height = 15 # Default height for the row
          # Check if row is not empty
          if not row or not hasattr(row[0], 'row'):
              continue

          row_idx = row[0].row # Get row index from the first cell

          for cell in row:
              if cell and cell.value:
                  try:
                      # Estimate height based on newlines, adjust as needed
                      lines = str(cell.value).count('\n') + 1
                      # Rough estimate: 15 points per line, add some padding
                      estimated_height = lines * 15 + (5 if lines > 1 else 0)
                      max_height = max(max_height, estimated_height)
                  except Exception as e:
                      # print(f"DEBUG: Error processing cell {cell.coordinate} for height: {e}")
                      pass

          # Apply calculated max height for the row, with a maximum limit
          worksheet.row_dimensions[row_idx].height = min(max_height, 150) # Max height 150


def post_process_rows(rows: List[List[str]]) -> List[List[str]]:
    """『合計』を含むセルの直上セルを空白にする処理"""
    if not rows: # Handle empty input
        return []
    # Create a deep copy to avoid modifying the original list structure
    # Ensure all elements are lists before copying
    processed_rows = [list(row) if isinstance(row, (list, tuple)) else [] for row in rows]

    for i, row in enumerate(processed_rows):
        # Ensure row is a list and not empty
        if not isinstance(row, list) or not row:
            continue
        for j, cell in enumerate(row):
            # Check if '合計' is present and it's not the first row
            # Also ensure the cell above exists and the row above is valid
            if "合計" in str(cell) and i > 0 and \
               isinstance(processed_rows[i-1], list) and j < len(processed_rows[i-1]):
                 processed_rows[i-1][j] = "" # Clear the cell directly above
    return processed_rows

def pdf_data_to_dataframe(pdf_file) -> pd.DataFrame | None:
    """
    PDFを読み込み、1ページ目のデータを抽出してDataFrameとして返す。
    """
    try:
        with pdfplumber.open(pdf_file) as pdf:
            if not pdf.pages:
                st.error("PDFファイルにページが含まれていません。")
                return None
            page = pdf.pages[0] # Process only the first page

            # Attempt layout extraction first
            rows = extract_text_with_layout(page)

            # Fallback to table extraction if layout fails or returns empty
            if not rows:
                 st.warning("レイアウト解析でデータを抽出できませんでした。テーブル抽出を試みます。")
                 tables = page.extract_tables()
                 if tables:
                     st.info("テーブル抽出成功。")
                     # Assuming the first table is the main one
                     # Clean table data: replace None with empty strings
                     rows = [[str(cell) if cell is not None else "" for cell in row] for row in tables[0]]
                 else:
                     st.error("テーブル抽出も失敗しました。")
                     return None

            # Post-processing and cleaning (applied to both layout and table data)
            rows = [row for row in rows if any(str(cell).strip() for cell in row if cell is not None)] # Remove fully empty rows
            if not rows:
                 st.error("有効なデータ行が見つかりませんでした。")
                 return None

            rows = post_process_rows(rows)
            rows = remove_extra_empty_columns(rows) # Crucial step after post_process

            if not rows:
                 st.error("クリーンアップ後、データが空になりました。")
                 return None

            # Find max columns again after all cleaning
            max_cols = max(len(row) for row in rows if row) if rows else 0
            if max_cols == 0:
                 st.error("最終的なデータ列数が0です。")
                 return None

            # Normalize rows to have the same number of columns for DataFrame creation
            normalized_rows = []
            for row in rows:
                 if row: # Ensure row is not None or empty
                     normalized_rows.append((row + [None] * (max_cols - len(row))) if len(row) < max_cols else row[:max_cols]) # Ensure correct length
                 else:
                     normalized_rows.append([None] * max_cols) # Add empty row placeholder


            # Create DataFrame without header
            df = pd.DataFrame(normalized_rows)
            return df

    except Exception as e:
        st.error(f"PDF処理中に予期せぬエラーが発生しました: {e}")
        # Log the full traceback for debugging if needed
        import traceback
        st.error(traceback.format_exc())
        return None


# ----------------------------
# ★ 新しい関数：シートの値をコピー ★
# ----------------------------
def copy_sheet_values(source_ws: Worksheet, target_ws: Worksheet):
    """
    source_ws のセルの値（数式ではなく結果）を target_ws にコピーする。
    target_ws の既存の内容はクリアされる。
    書式はコピーしない。列幅/行高さは別途調整。
    """
    if not isinstance(source_ws, Worksheet) or not isinstance(target_ws, Worksheet):
        st.error("コピー元またはコピー先のシートが無効です。")
        return

    # ターゲットシートをクリア
    # iter_rows(max_row=...) can be unreliable if rows were deleted without updating max_row
    # Safest is to delete all rows
    try:
         if target_ws.max_row > 0: # Check if there are rows to delete
             target_ws.delete_rows(1, target_ws.max_row + 1) # Clear all existing rows
    except Exception as e:
         st.warning(f"ターゲットシート '{target_ws.title}' のクリア中にエラー: {e}. 処理を続行します。")


    # ソースシートから値をコピー
    try:
        for r_idx, row in enumerate(source_ws.iter_rows(), 1):
            for c_idx, cell in enumerate(row, 1):
                # value属性を使って値のみを取得
                 target_ws.cell(row=r_idx, column=c_idx, value=cell.value)
    except Exception as e:
         st.error(f"シート '{source_ws.title}' から '{target_ws.title}' へのコピー中にエラー: {e}")
         import traceback
         st.error(traceback.format_exc())
         raise # Re-raise the exception to stop the process if copy fails critically


    # 列幅と行高さを調整 (コピー後に実行)
    try:
        format_excel_worksheet(target_ws)
    except Exception as e:
         st.warning(f"シート '{target_ws.title}' のフォーマット中にエラー: {e}")


# ----------------------------
# Excelファイルのパス設定と存在確認
# ----------------------------
template_path = "template.xlsx"
release_path = "release.xlsx"
template_wb = None
release_wb = None
error_messages = []

if not os.path.exists(template_path):
    error_messages.append(f"テンプレートファイル '{template_path}' が見つかりません。")
else:
    try:
        # data_only=True で数式の代わりに値を読み込む (ただし、template側で必要ならFalse)
        template_wb = load_workbook(template_path, data_only=False)
    except Exception as e:
        error_messages.append(f"テンプレートファイル '{template_path}' の読み込みに失敗しました: {e}")

if not os.path.exists(release_path):
    error_messages.append(f"リリース用ファイル '{release_path}' が見つかりません。")
else:
    try:
        release_wb = load_workbook(release_path)
    except Exception as e:
        error_messages.append(f"リリース用ファイル '{release_path}' の読み込みに失敗しました: {e}")

# エラーがあれば表示して停止
if error_messages:
    for msg in error_messages:
        st.error(msg)
    st.stop()

# ----------------------------
# UI：PDFファイルアップロード＆変換実行
# ----------------------------
uploaded_pdf = st.file_uploader("", type="pdf",
                                help="PDFをアップロードするとExcelに変換され、テンプレートの1シート目に貼り付け、その後指定シートが別ファイルにコピーされます")

file_container = st.container()
processed = False # 処理状態フラグ

if uploaded_pdf and template_wb and release_wb: # 両方のWorkbookが正常に読み込めた場合のみ実行
    file_ext = uploaded_pdf.name.split('.')[-1].lower()
    file_icon = "PDF" if file_ext == "pdf" else file_ext.upper()
    # Ensure getvalue() returns bytes before calculating size
    pdf_bytes = uploaded_pdf.getvalue()
    file_size = len(pdf_bytes) / 1024 if pdf_bytes else 0 # KB単位

    # --- ファイル情報表示（処理中）---
    progress_placeholder = st.empty() # プレースホルダーを作成
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
    <div class="progress-bar"><div class="progress-value" id="progress-bar-value"></div></div>
    """, unsafe_allow_html=True)
    # --- ここまで ---

    final_excel_bytes = None # ダウンロード用バイト列を初期化

    with st.spinner("ファイル処理中..."):
        try:
            # 1. PDFからDataFrameへ変換
            st.write("ステップ1/3: PDFからデータを抽出中...")
            # Pass the bytes directly to the function
            df_pdf = pdf_data_to_dataframe(io.BytesIO(pdf_bytes))

            if df_pdf is not None and not df_pdf.empty:
                # 2. DataFrameをtemplate.xlsxの1シート目に書き込み
                st.write("ステップ2/3: テンプレートファイルにデータを書き込み中...")
                try:
                    if not template_wb.worksheets:
                         st.error(f"'{template_path}' にシートが存在しません。")
                         st.stop()

                    template_ws_target = template_wb.worksheets[0] # 最初のシートを取得

                    # 既存の内容をクリア（オプション） - より安全な方法
                    if template_ws_target.max_row > 0:
                        template_ws_target.delete_rows(1, template_ws_target.max_row + 1)

                    # DataFrameを書き込み (ヘッダーなし、インデックスなし)
                    for r_idx, row in enumerate(df_pdf.values):
                        for c_idx, value in enumerate(row):
                             # NaN値をNoneに変換 (openpyxlはNaNを扱えないため)
                             if pd.isna(value):
                                 value = None
                             # Ensure value is suitable for Excel cell
                             if isinstance(value, (list, tuple, dict)):
                                 value = str(value) # Convert complex types to string
                             template_ws_target.cell(row=r_idx + 1, column=c_idx + 1, value=value)

                    format_excel_worksheet(template_ws_target) # 書式設定

                except IndexError:
                    # This should be caught by the check above, but keep as safeguard
                    st.error(f"'{template_path}' に最初のシートが見つかりません。")
                    st.stop()
                except Exception as e:
                    st.error(f"テンプレートへの書き込み中にエラーが発生しました: {e}")
                    import traceback
                    st.error(traceback.format_exc())
                    st.stop()


                # 3. template.xlsxからrelease.xlsxへ値をコピー
                st.write("ステップ3/3: リリースファイルへデータをコピー中...")
                # --- コピー元・コピー先シート名の指定 ---
                source_sheet_name_1 = "数出表_Excel（アレルギー入力）" # templateの3シート目相当
                source_sheet_name_2 = "盛付札"                 # templateの4シート目相当
                target_sheet_index_1 = 0                      # releaseの1シート目
                target_sheet_index_2 = 1                      # releaseの2シート目

                copy_successful = True # コピー成功フラグ
                source_ws1 = None
                source_ws2 = None

                # --- Find Source Sheet 1 ---
                if source_sheet_name_1 in template_wb.sheetnames:
                    source_ws1 = template_wb[source_sheet_name_1]
                elif len(template_wb.worksheets) > 2:
                     source_ws1 = template_wb.worksheets[2] # Index 2 is the 3rd sheet
                     st.warning(f"シート名 '{source_sheet_name_1}' が見つかりません。3番目のシートを使用します。")
                else:
                     st.error(f"シート '{source_sheet_name_1}' も3番目のシートも '{template_path}' に存在しません。")
                     copy_successful = False

                # --- Find Source Sheet 2 ---
                if copy_successful: # Only proceed if Sheet 1 was found
                    if source_sheet_name_2 in template_wb.sheetnames:
                        source_ws2 = template_wb[source_sheet_name_2]
                    elif len(template_wb.worksheets) > 3:
                        source_ws2 = template_wb.worksheets[3] # Index 3 is the 4th sheet
                        st.warning(f"シート名 '{source_sheet_name_2}' が見つかりません。4番目のシートを使用します。")
                    else:
                        st.error(f"シート '{source_sheet_name_2}' も4番目のシートも '{template_path}' に存在しません。")
                        copy_successful = False


                # --- Find Target Sheets & Perform Copy ---
                target_ws1 = None
                target_ws2 = None

                if copy_successful:
                    if len(release_wb.worksheets) > target_sheet_index_1:
                        target_ws1 = release_wb.worksheets[target_sheet_index_1]
                    else:
                        st.error(f"'{release_path}' に {target_sheet_index_1 + 1}番目のシートが存在しません。")
                        copy_successful = False

                if copy_successful:
                     if len(release_wb.worksheets) > target_sheet_index_2:
                         target_ws2 = release_wb.worksheets[target_sheet_index_2]
                     else:
                         st.error(f"'{release_path}' に {target_sheet_index_2 + 1}番目のシートが存在しません。")
                         copy_successful = False

                # --- Execute Copying ---
                if copy_successful and source_ws1 and target_ws1:
                    try:
                        copy_sheet_values(source_ws1, target_ws1)
                    except Exception as e:
                        st.error(f"シート '{source_ws1.title}' から '{target_ws1.title}' へのコピー中にエラー: {e}")
                        copy_successful = False

                if copy_successful and source_ws2 and target_ws2:
                     try:
                         copy_sheet_values(source_ws2, target_ws2)
                     except Exception as e:
                         st.error(f"シート '{source_ws2.title}' から '{target_ws2.title}' へのコピー中にエラー: {e}")
                         copy_successful = False


                # 4. 最終的なrelease.xlsxをバイトデータとして保存
                if copy_successful:
                    try:
                        output_release = io.BytesIO()
                        release_wb.save(output_release)
                        output_release.seek(0)
                        final_excel_bytes = output_release.read()
                        processed = True # すべて成功した場合にフラグを立てる
                    except Exception as e:
                        st.error(f"最終Excelファイルの保存中にエラーが発生しました: {e}")
                        processed = False # エラー発生
                else:
                     # Ensure processed is False if copy failed at any point
                     processed = False

            else:
                # PDFからのデータ抽出失敗
                st.error("PDFからのデータ抽出に失敗したため、処理を中断しました。")
                processed = False

        except Exception as e:
            st.error(f"予期せぬエラーが発生しました: {e}")
            import traceback
            st.error(traceback.format_exc()) # 詳細なエラーログ
            processed = False

    # --- ファイル情報表示（処理完了 or エラー）---
    # Make sure progress_placeholder exists before updating
    if 'progress_placeholder' in locals():
        if processed:
            # 成功時の表示
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
             <div class="progress-bar"><div class="progress-value done"></div></div>
            """, unsafe_allow_html=True)
        else:
            # 失敗時の表示
             progress_placeholder.markdown(f"""
             <div class="file-card" style="border-color: #f44336;">
                 <div class="file-info">
                     <div class="file-icon">{file_icon}</div>
                     <div class="file-details">
                         <div class="file-name">{uploaded_pdf.name}</div>
                         <div class="file-meta">{file_size:.0f} KB - <span style="color: #f44336; font-weight: bold;">処理失敗</span></div>
                     </div>
                 </div>
                 <div style="color: #f44336; font-size: 20px; font-weight: bold;">✕</div>
             </div>
             """, unsafe_allow_html=True)
    # --- ここまで ---


    # 処理が成功し、ダウンロード用データがある場合のみダウンロードリンクを表示
    if processed and final_excel_bytes:
        st.markdown('<div class="separator"></div>', unsafe_allow_html=True)

        original_pdf_name = os.path.splitext(uploaded_pdf.name)[0]
        # ★ 出力ファイル名を release.xlsx ベースに変更（必要なら調整）★
        # output_filename = f"{original_pdf_name}_Release.xlsx"
        output_filename = f"{original_pdf_name}_Merged.xlsx" # 元の命名規則を維持

        excel_size = len(final_excel_bytes) / 1024 # KB単位
        b64 = base64.b64encode(final_excel_bytes).decode('utf-8')

        # ダウンロードリンクのHTML (クラス名を修正、コメント削除)
        href = f"""
        <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{output_filename}" class="download-card">
            <div class="download-info">
                <div class="excel-icon">XLSX</div>
                <div class="download-details">
                    <div class="download-name">{output_filename}</div>
                    <div class="download-meta">変換済みExcel・{excel_size:.0f} KB</div>
                </div>
            </div>
            <button class="download-button">
                <span class="download-button-icon">↓</span>
                Download
            </button>
        </a>
        """
        st.markdown(href, unsafe_allow_html=True)
    # elif not processed and uploaded_pdf: # Redundant check, error shown above
    #      st.warning("処理中にエラーが発生したため、ファイルをダウンロードできません。")


# メインコンテナ終了
st.markdown('</div>', unsafe_allow_html=True)
