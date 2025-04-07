import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import base64
import os
from typing import List, Dict, Any
from openpyxl import load_workbook

# (他の部分は省略)

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
        # (中略: PDF->Excel変換、テンプレート書き込み処理)
        converted_excel_bytes = pdf_to_excel(uploaded_pdf)
        if converted_excel_bytes is None:
            st.error("PDFからデータを抽出できませんでした。")
            st.stop()
        df_pdf = pd.read_excel(io.BytesIO(converted_excel_bytes), sheet_name=0, header=None)

        template_wb = load_workbook(template_path, keep_vba=True) # 再度読み込むか、前の処理で保持
        template_ws = template_wb.worksheets[0]

        for r_idx, row in df_pdf.iterrows():
            for c_idx, value in enumerate(row):
                template_ws.cell(row=r_idx+1, column=c_idx+1, value=value)

        output = io.BytesIO()
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
    output_filename = f"{original_pdf_name}_Merged.xlsm"
    excel_size = len(final_excel_bytes) / 1024  # KB単位
    b64 = base64.b64encode(final_excel_bytes).decode('utf-8')
    mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12"

    # --- 修正箇所: HTMLコメントを削除 ---
    href = f"""
    <a href="data:{mime_type};base64,{b64}" download="{output_filename}" class="download-card">
        <div class="download-info">
            <div class="download-icon">XLSM</div> {/* コメント削除 */}
            <div class="download-details">
                <div class="download-name">{output_filename}</div>
                <div class="download-meta">Excel (マクロ有効)・{excel_size:.0f} KB</div> {/* コメント削除 */}
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
