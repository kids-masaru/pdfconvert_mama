import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import base64
import os
from typing import List, Dict, Any
from openpyxl import load_workbook

# ----------------------------
# ページ設定（アイコン指定：ブラウザタブ・ブックマーク用）
# ----------------------------
st.set_page_config(
    page_title="【数出表】PDF → Excelへの変換",
    layout="centered",
    page_icon="icon.png"  # アイコンファイルのパスを指定
)

# ----------------------------
# UIのスタイル設定（洗練されたモダンデザイン - 暖色系背景）
# ----------------------------
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
            background-color: #f44336;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 12px;
            color: white;
            font-weight: 500;
            font-size: 14px;
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
            border-top-color: #ff9933;
            animation: spin 1s linear infinite;
        }

        .check-icon {
            color: #ff9933;
            font-size: 20px;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        /* 進行状況バー */
        .progress-bar {
            height: 4px;
            background-color: #e0e0e0;
            border-radius: 2px;
            width: 100%;
            margin-top: 10px;
        }

        .progress-value {
            height: 100%;
            background-color: #ff9933;
            border-radius: 2px;
            width: 60%;
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
            text-decoration: none; /* 下線を削除 */
        }

        .download-card:hover {
            box-shadow: 0 4px 12px rgba(0,0,0,0.12);
            background-color: #fffaf0;
            transform: translateY(-2px);
        }

        .download-info {
            display: flex;
            align-items: center;
        }

        .download-icon {
            width: 40px;
            height: 40px;
            border-radius: 8px;
            background-color: #ff9933;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 12px;
            color: white;
            font-weight: 500;
            font-size: 16px;
        }

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
            background-color: #ff9933;
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
            background-color: #e68a00;
        }

        .download-button-icon {
            margin-right: 6px;
        }

        .stSpinner > div {
            display: none;
        }

        .css-1544g2n {
            padding-top: 2rem;
        }

        .css-18e3th9 {
            padding-top: 2rem;
        }

        /* セパレーター */
        .separator {
            height: 1px;
            background-color: #ffe0b3;
            margin: 25px 0;
        }
    </style>
""", unsafe_allow_html=True)

# メインコンテナ開始
st.markdown('<div class="main-container">', unsafe_allow_html=True)

# タイトルとサブタイトル（メインコンテナ内に配置）
st.markdown('<div class="title">【数出表】PDF → Excelへの変換</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">PDFの数出表をExcelに変換し、同時に盛り付け札を作成します。</div>', unsafe_allow_html=True)

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
    """PDFのページ内から、縦線とみなせる線のx座標を抽出する"""
    vertical_lines = []
    for line in page.lines:
        if abs(line['x0'] - line['x1']) < tolerance:
            vertical_lines.append((line['x0'] + line['x1']) / 2)
    vertical_lines = sorted(set(round(x, 1) for x in vertical_lines))
    words = page.extract_words()
    if not words:
        return vertical_lines
    left_boundary = min(word['x0'] for word in words)
    right_boundary = max(word['x1'] for word in words)
    boundaries = [left_boundary] + vertical_lines + [right_boundary]
    boundaries = sorted(boundaries)
    return boundaries

def split_line_using_boundaries(sorted_words: List[Dict[str, Any]], boundaries: List[float]) -> List[str]:
    """同一行内の単語をセルごとに分割する"""
    columns = []
    for i in range(len(boundaries) - 1):
        left = boundaries[i]
        right = boundaries[i+1]
        col_words = [word['text'] for word in sorted_words
                     if (word['x0'] + word['x1'])/2 >= left and (word['x0'] + word['x1'])/2 < right]
        cell_text = " ".join(col_words)
        columns.append(cell_text)
    return columns

def extract_text_with_layout(page) -> List[List[str]]:
    """PDFページからセル分割されたテキストデータを抽出する"""
    words = page.extract_words(x_tolerance=5, y_tolerance=5)
    if not words:
        return []
    boundaries = get_vertical_boundaries(page)
    row_groups = get_line_groups(words, y_tolerance=1.2)
    result_rows = []
    for group in row_groups:
        sorted_group = sorted(group, key=lambda w: w['x0'])
        if boundaries:
            columns = split_line_using_boundaries(sorted_group, boundaries)
        else:
            columns = [" ".join(word['text'] for word in sorted_group)]
        result_rows.append(columns)
    return result_rows

def remove_extra_empty_columns(rows: List[List[str]]) -> List[List[str]]:
    """すべての行で空の列を削除する"""
    if not rows:
        return rows
    num_cols = max(len(row) for row in rows)
    keep_indices = []
    for col in range(num_cols):
        if any(row[col].strip() for row in rows if col < len(row)):
            keep_indices.append(col)
    new_rows = []
    for row in rows:
        new_row = [row[i] if i < len(row) else "" for i in keep_indices]
        new_rows.append(new_row)
    return new_rows

def format_excel_worksheet(worksheet):
    """Excelワークシートの書式設定（列幅・行高さ）"""
    # 注意: xlsxwriter の機能なので、openpyxl でロードしたワークブックには直接適用できません。
    # openpyxl で書式設定を行う場合は、別途 openpyxl のメソッドを使用する必要があります。
    # この関数は pdf_to_excel 関数内で xlsxwriter を使っている部分でのみ有効です。
    try:
        worksheet.set_column('A:Z', 20)
        worksheet.set_default_row(20)
    except AttributeError:
        # openpyxl のワークシートオブジェクトにはこれらのメソッドがないため、エラーを無視
        pass


def post_process_rows(rows: List[List[str]]) -> List[List[str]]:
    """『合計』を含むセルの直上セルを空白にする処理"""
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            if "合計" in cell:
                if i > 0 and j < len(rows[i-1]):
                    rows[i-1][j] = ""
    return rows

def pdf_to_excel(pdf_file):
    """
    PDFを読み込み、1ページ目のデータをExcelシートとして出力する。
    ※変換後のExcelはタブ（シート）が1つになる前提です。
    xlsxwriter を使用して一時的な Excel ファイルをメモリ上に作成します。
    """
    output = io.BytesIO()
    with pdfplumber.open(pdf_file) as pdf:
        # 一時的なExcelファイル作成には xlsxwriter を使用
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            border_format = workbook.add_format({'border': 1})
            page = pdf.pages[0]
            rows = extract_text_with_layout(page)
            rows = [row for row in rows if any(cell.strip() for cell in row)]
            if not rows:
                return None
            rows = post_process_rows(rows)
            rows = remove_extra_empty_columns(rows)
            max_cols = max(len(row) for row in rows)
            normalized_rows = [row + [''] * (max_cols - len(row)) for row in rows]
            df = pd.DataFrame(normalized_rows)
            sheet_name = "ConvertedData"
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            worksheet = writer.sheets[sheet_name]
            format_excel_worksheet(worksheet) # xlsxwriter の worksheet オブジェクトを渡す
            # 罫線の適用
            for r in range(len(normalized_rows)):
                for c in range(max_cols):
                    worksheet.write(r, c, normalized_rows[r][c], border_format)
    output.seek(0)
    return output.read()

# ----------------------------
# テンプレートExcelファイルの自動読み込み
# ----------------------------
# --- 修正点 1: ファイルパスを変更 ---
template_path = "template.xlsm"
if not os.path.exists(template_path):
    st.error(f"テンプレートファイル '{template_path}' がサーバー上に存在しません。")
    st.stop()
else:
    try:
        # --- 修正点 2: keep_vba=True を追加 ---
        template_wb = load_workbook(template_path, keep_vba=True)
    except Exception as e:
        st.error(f"テンプレートファイル '{template_path}' の読み込み中にエラーが発生しました: {e}")
        st.stop()

# ----------------------------
# UI：PDFファイルアップロード＆変換実行
# ----------------------------
uploaded_pdf = st.file_uploader("", type="pdf",
                                help="PDFをアップロードするとExcelに変換され、テンプレートの1シート目に貼り付けます")

file_container = st.container()
processed = False

if uploaded_pdf:
    file_ext = uploaded_pdf.name.split('.')[-1].lower()
    file_icon = "PDF" if file_ext == "pdf" else file_ext.upper()
    file_size = len(uploaded_pdf.getvalue()) / 1024  # KB単位

    with file_container:
        if not processed:
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
            <div class="progress-bar">
                <div class="progress-value"></div>
            </div>
            """, unsafe_allow_html=True)

    with st.spinner("変換中..."):
        converted_excel_bytes = pdf_to_excel(uploaded_pdf)
        if converted_excel_bytes is None:
            st.error("PDFからデータを抽出できませんでした。")
            st.stop()
        # pdf_to_excel で作成された一時Excel（メモリ上）を読み込む
        df_pdf = pd.read_excel(io.BytesIO(converted_excel_bytes), sheet_name=0, header=None)

        # template.xlsm の最初のシートを取得
        template_ws = template_wb.worksheets[0]

        # データをテンプレートシートに書き込む
        # ※ 既存のデータをクリアしたい場合は、ここでクリア処理を追加します。
        # 例: template_ws.delete_rows(1, template_ws.max_row) など
        for r_idx, row in df_pdf.iterrows():
            for c_idx, value in enumerate(row):
                # openpyxlは1始まりのインデックス
                template_ws.cell(row=r_idx+1, column=c_idx+1, value=value)

        # 変更をメモリ上の BytesIO オブジェクトに保存
        output = io.BytesIO()
        # --- template_wb は keep_vba=True でロードされているため、マクロは保持される ---
        template_wb.save(output)
        output.seek(0)
        final_excel_bytes = output.read()

        processed = True

    with file_container:
        if processed:
            progress_placeholder.markdown(f"""
            <div class="file-card">
                <div class="file-info">
                    <div class="file-icon">{file_icon}</div>
                    <div class="file-details">
                        <div class="file-name">{uploaded_pdf.name}</div>
                        <div class="file-meta">{file_size:.0f} KB</div>
                    </div>
                </div>
                <div class="check-icon">✓</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown('<div class="separator"></div>', unsafe_allow_html=True)

    original_pdf_name = os.path.splitext(uploaded_pdf.name)[0]
    # --- 修正点 3: 出力ファイル名を .xlsm に変更 ---
    output_filename = f"{original_pdf_name}_Merged.xlsm"
    excel_size = len(final_excel_bytes) / 1024  # KB単位
    b64 = base64.b64encode(final_excel_bytes).decode('utf-8')

    # --- 修正点 4: MIMEタイプを .xlsm 用に変更 ---
    mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12"
    href = f"""
    <a href="data:{mime_type};base64,{b64}" download="{output_filename}" class="download-card">
        <div class="download-info">
            <div class="download-icon">XLSM</div> {/* アイコンも変更 */}
            <div class="download-details">
                <div class="download-name">{output_filename}</div>
                <div class="download-meta">Excel (マクロ有効)・{excel_size:.0f} KB</div> {/* メタ情報も変更 */}
            </div>
        </div>
        <button class="download-button">
            <span class="download-button-icon">↓</span>
            Download
        </button>
    </a>
    """
    st.markdown(href, unsafe_allow_html=True)

# メインコンテナ終了
st.markdown('</div>', unsafe_allow_html=True)
