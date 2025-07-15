import streamlit as st
import pdfplumber
import pandas as pd
import io

# ページ設定
st.set_page_config(
    page_title="PDF Convert Mama - PDFをExcelに変換",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="auto"
)

# サイドバーのタイトル
st.sidebar.title("メニュー")

# サイドバーに他のページへのリンクを配置
# マスタ設定ページへのリンクはそのまま残します
st.sidebar.page_link("pages/2_Master_Data_Settings.py", label="マスタ設定", icon="⚙️")

# --- メインコンテンツ（PDF → Excel 変換機能） ---
st.title("📄 PDF を Excel に変換")
st.markdown("---")

st.write("PDFファイルをアップロードして、データをExcel形式に変換します。")

uploaded_file = st.file_uploader("PDFファイルをここにドラッグ＆ドロップ、またはクリックしてアップロード", type="pdf")

if uploaded_file is not None:
    try:
        # PDFを読み込み
        with pdfplumber.open(uploaded_file) as pdf:
            all_text = ""
            for page in pdf.pages:
                all_text += page.extract_text() + "\n" # 各ページのテキストを抽出

            # ここから、抽出したテキストをDataFrameに変換するロジックを実装します。
            # 例: テキストを行ごとに分割し、DataFrameを作成
            # ※ 実際のPDF構造に応じて、より複雑な抽出ロジックが必要です
            lines = all_text.split('\n')
            
            # 例として、単純に各行を1つの列としてDataFrameにする
            # 実際のデータ抽出ロジックは、PDFの内容に応じてこの部分を詳細に記述してください
            if lines:
                # 空行を除去し、各行をリストの要素とする
                cleaned_lines = [line.strip() for line in lines if line.strip()]
                if cleaned_lines:
                    df = pd.DataFrame(cleaned_lines, columns=["抽出テキスト"])
                    st.write("PDFから以下のテキストを抽出しました:")
                    st.dataframe(df)

                    # DataFrameをExcelとしてダウンロード
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
