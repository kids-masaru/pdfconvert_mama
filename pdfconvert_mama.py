# pdfconvert_mama.py
import streamlit as st
import streamlit.components.v1 as components
import pdfplumber
import pandas as pd
import io
import re
import base64
import os
from typing import List, Dict, Any
from openpyxl import load_workbook  # .xlsm 読み書きのため

# ──────────────────────────────────────────────
# ① HTML <head> 埋め込み（PWA用 manifest & 各種アイコン）
# ──────────────────────────────────────────────
components.html(
    """
    <!-- PWA 用マニフェスト -->
    <link rel="manifest" href="/app/static/manifest.json">
    <!-- ブラウザタブ用 favicon -->
    <link rel="icon" href="/app/static/favicon.ico" type="image/x-icon">
    <!-- iOS ホーム画面用アイコン -->
    <link rel="apple-touch-icon" sizes="180x180" href="/app/static/icons/apple-touch-icon.png">
    <!-- iOS でネイティブ風表示 -->
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-title" content="数出表PDF→Excel">
    """,
    height=0,
)

# ──────────────────────────────────────────────
# ② Streamlit ページ設定（ブラウザタブのアイコンも static 配信を使う）
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="【数出表】PDF → Excelへの変換",
    page_icon="/app/static/favicon.ico",
    layout="centered",
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
        .file-icon { width: 36px; height: 36px; border-radius: 6px; background-color: #f44336; display: flex; align-items: center; justify-content: center; margin-right: 12px; color: white; font-weight: 500; font-size: 14px; }
        .loading-spinner { width: 20px; height: 20px; border: 2px solid rgba(0,0,0,0.1); border-radius: 50%; border-top-color: #ff9933; animation: spin 1s linear infinite; }
        .check-icon { color: #ff9933; font-size: 20px; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .progress-bar { height: 4px; background-color: #e0e0e0; border-radius: 2px; width: 100%; margin-top: 10px; overflow: hidden; }
        .progress-value { height: 100%; background-color: #ff9933; border-radius: 2px; width: 60%; transition: width 0.5s ease-in-out; }
        .download-card { background: white; border-radius: 8px; padding: 16px; margin: 20px 0; box-shadow: 0 2px 5px rgba(0,0,0,0.08); display: flex; align-items: center; justify-content: space-between; border: 1px solid #eaeaea; transition: all 0.2s ease; cursor: pointer; text-decoration: none; color: inherit; }
        .download-card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.12); background-color: #fffaf0; transform: translateY(-2px); }
        .separator { height: 1px; background-color: #ffe0b3; margin: 25px 0; }
    </style>
""", unsafe_allow_html=True)

# メインコンテナ開始
st.markdown('<div class="main-container">', unsafe_allow_html=True)
st.markdown('<div class="title">【数出表】PDF → Excelへの変換</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">PDFの数出表をExcelに変換し、同時に盛り付け札を作成します。</div>', unsafe_allow_html=True)

# ──────────────────────────────────────────────
# 以下、PDF→Excel変換ロジック（既存のまま貼り付け）
# ※関数定義は省略せずにこのまま入れてください。
# ──────────────────────────────────────────────
def is_number(text: str) -> bool:
    return bool(re.match(r'^\d+$', text.strip()))

def get_line_groups(words: List[Dict[str, Any]], y_tolerance: float = 1.2) -> List[List[Dict[str, Any]]]:
    if not words: return []
    sorted_words = sorted(words, key=lambda w: w['top'])
    groups, current_group, current_top = [], [sorted_words[0]], sorted_words[0]['top']
    for w in sorted_words[1:]:
        if abs(w['top'] - current_top) <= y_tolerance:
            current_group.append(w)
        else:
            groups.append(current_group)
            current_group, current_top = [w], w['top']
    groups.append(current_group)
    return groups

def get_vertical_boundaries(page, tolerance: float = 2) -> List[float]:
    xs = [(l['x0']+l['x1'])/2 for l in page.lines if abs(l['x0']-l['x1'])<tolerance]
    xs = sorted(set(round(x,1) for x in xs))
    words = page.extract_words()
    if not words: return xs
    left = min(w['x0'] for w in words); right = max(w['x1'] for w in words)
    b = sorted(set([round(left,1)] + xs + [round(right,1)]))
    merged = [b[0]]
    for x in b[1:]:
        if x-merged[-1] > tolerance*2: merged.append(x)
    if right > merged[-1]+tolerance*2: merged.append(round(right,1))
    return sorted(set(merged))

def split_line_using_boundaries(sorted_words, boundaries) -> List[str]:
    cols = [""]*(len(boundaries)-1)
    for w in sorted_words:
        cx = (w['x0']+w['x1'])/2
        for i in range(len(boundaries)-1):
            if boundaries[i] <= cx < boundaries[i+1]:
                cols[i] += w['text'] + " "
                break
    return [c.strip() for c in cols]

def extract_text_with_layout(page) -> List[List[str]]:
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    if not words: return []
    bd = get_vertical_boundaries(page)
    if len(bd)<2:
        text = page.extract_text(layout=False, x_tolerance=3, y_tolerance=3)
        return [[l] for l in text.split("\n") if l.strip()]
    rows = []
    for grp in get_line_groups(words, 1.5):
        grp_sorted = sorted(grp, key=lambda w: w['x0'])
        row = split_line_using_boundaries(grp_sorted, bd)
        if any(c.strip() for c in row): rows.append(row)
    return rows

def remove_extra_empty_columns(rows: List[List[str]]) -> List[List[str]]:
    if not rows: return []
    ncol = max(len(r) for r in rows)
    empty = [True]*ncol
    for r in rows:
        for i,c in enumerate(r):
            if c.strip(): empty[i]=False
    keep = [i for i,e in enumerate(empty) if not e]
    return [[r[i] if i<len(r) else "" for i in keep] for r in rows]

def post_process_rows(rows: List[List[str]]) -> List[List[str]]:
    nr = [r[:] for r in rows]
    for i,row in enumerate(nr):
        for j,cell in enumerate(row):
            if "合計" in str(cell) and i>0 and j<len(nr[i-1]):
                nr[i-1][j] = ""
    return nr

def pdf_to_excel_data(f) -> pd.DataFrame| None:
    try:
        with pdfplumber.open(f) as pdf:
            if not pdf.pages: st.warning("PDFにページがありません。"); return None
            rows = extract_text_with_layout(pdf.pages[0])
            rows = [r for r in rows if any(c.strip() for c in r)]
            if not rows: st.warning("テキスト抽出できず。"); return None
            rows = post_process_rows(rows)
            rows = remove_extra_empty_columns(rows)
            if not rows or not rows[0]: st.warning("データ消失。"); return None
            mx = max(len(r) for r in rows)
            norm = [r+[""]*(mx-len(r)) for r in rows]
            return pd.DataFrame(norm)
    except Exception as e:
        st.error(f"PDF処理エラー: {e}"); return None

# テンプレート読み込み
template_path = "template.xlsm"
if not os.path.exists(template_path):
    st.error(f"テンプレート'{template_path}'が見つかりません。"); st.stop()
try:
    template_wb = load_workbook(template_path, keep_vba=True)
except Exception as e:
    st.error(f"テンプレート読み込み失敗: {e}"); st.stop()

# UI：アップローダー
uploaded = st.file_uploader("PDFをアップロード", type="pdf")
file_c = st.container(); dl_c = st.container()

if uploaded and template_wb:
    # 処理中UI
    with file_c:
        sz = len(uploaded.getvalue())/1024
        file_c.markdown(f"""
            <div class="file-card">
                <div class="file-icon">PDF</div>
                <div class="file-details">
                    <div class="file-name">{uploaded.name}</div>
                    <div class="file-meta">{sz:.0f} KB 処理中…</div>
                </div>
                <div class="loading-spinner"></div>
            </div>
        """, unsafe_allow_html=True)

    df = pdf_to_excel_data(uploaded)
    if df is not None and not df.empty:
        # テンプレートに書き込み
        ws = template_wb.worksheets[0]
        for i,row in df.iterrows():
            for j,val in enumerate(row):
                ws.cell(row=i+1, column=j+1, value=val)

        # 出力
        buf = io.BytesIO()
        template_wb.save(buf); buf.seek(0)
        data = buf.read()
        b64 = base64.b64encode(data).decode()
        fname = os.path.splitext(uploaded.name)[0] + "_Merged.xlsm"
        with dl_c:
            st.markdown('<div class="separator"></div>', unsafe_allow_html=True)
            st.markdown(f'''
                <a href="data:application/vnd.ms-excel.sheet.macroEnabled.12;base64,{b64}"
                   download="{fname}" class="download-card">
                  <div class="download-details">
                    <div class="download-name">{fname}</div>
                    <div class="download-meta">{len(data)/1024:.0f} KB</div>
                  </div>
                  <div class="download-button-imitation">
                    ↓ Download
                  </div>
                </a>
            ''', unsafe_allow_html=True)
    else:
        with file_c:
            st.warning("PDFからデータを取得できませんでした。")
