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

# --- サイドバーナビゲーション ---
st.sidebar.title("メニュー")
page_selection = st.sidebar.radio(
    "表示する機能を選択してください",
    ("PDF → Excel 変換", "マスタ設定"),
    index=0 # 初期表示は「PDF → Excel 変換」
)

st.markdown("---") # メインコンテンツとサイドバーの区切り

# --- メインコンテンツの表示ロジック ---

# PDF → Excel 変換 ページ
if page_selection == "PDF → Excel 変換":
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

# マスタ設定 ページ (以前私が提供した仮の内容を統合)
elif page_selection == "マスタ設定":
    st.title("⚙️ マスタデータ設定")
    st.write("ここにマスタデータ設定の具体的な機能（例：CSVアップロード、データ表示、編集、保存など）を実装します。")

    # 以前私が提供した仮のマスターデータ表示のコード
    st.subheader("現在のマスタデータ（仮）")
    st.dataframe(pd.DataFrame({"ID": [1, 2, 3], "Name": ["Item A", "Item B", "Item C"]}))

    st.info("※ このマスタデータ設定は仮の機能です。実際の用途に合わせて、コードを追記・修正してください。")
