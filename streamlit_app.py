import streamlit as st
import pdfplumber
import pandas as pd
import io

# --- ページ設定 ---
st.set_page_config(
    page_title="PDF Convert Mama",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="auto"
)

# --- サイドバーのタイトル ---
st.sidebar.title("メニュー")

# --- サイドバーに他のページへのリンクを配置（ページへのリンクは残す）---
# Home ページへのリンクは不要なので削除
st.sidebar.page_link("pages/2_Master_Data_Settings.py", label="マスタ設定", icon="⚙️")
# もし今後、他のページを作成するならここに追加します

# --- メインコンテンツ（PDF → Excel 変換機能） ---
# ラジオボタンによるページ選択ロジックは全て削除します
# 以下は、PDF変換機能のコードのみ残します
st.title("📄 PDF を Excel に変換")
st.write("PDFファイルをアップロードして、データをExcel形式に変換します。")

uploaded_file = st.file_uploader("PDFファイルをここにドラッグ＆ドロップ、またはクリックしてアップロード", type="pdf")

if uploaded_file is not None:
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            all_text = ""
            for page in pdf.pages:
                all_text += page.extract_text() + "\n"

            lines = all_text.split('\n')
            if lines:
                cleaned_lines = [line.strip() for line in lines if line.strip()]
                if cleaned_lines:
                    df = pd.DataFrame(cleaned_lines, columns=["抽出テキスト"])
                    st.write("PDFから以下のテキストを抽出しました:")
                    st.dataframe(df)

                    excel_data = io.BytesIO()
                    with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='Sheet1')

                    st.download_button(
                        label="Excelファイルをダウンロード",
                        data=excel_data.getvalue(),
                        file_name="extracted_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("PDFから抽出できるテキストが見つかりませんでした。")
            else:
                st.warning("PDFからテキストを抽出できませんでした。")

    except Exception as e:
        st.error(f"ファイルの処理中にエラーが発生しました: {e}")
        st.info("PDFの内容やフォーマットが原因である可能性があります。別のPDFでお試しください。")

st.markdown("---")
st.info("※ このアプリはデモ目的で、PDFからのテキスト抽出を簡略化しています。")
st.info("実際のPDFファイルからのデータ抽出には、`pdfplumber`のテーブル抽出機能や正規表現など、より高度なロジックが必要です。")

# `elif page_selection == "マスタ設定":` から始まるマスタ設定のコードブロックも全て削除します。
# マスタ設定のコードは `pages/2_Master_Data_Settings.py` に残します。
