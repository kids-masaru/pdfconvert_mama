import streamlit as st
import pandas as pd
import os
import io

st.markdown('<div class="main-container">', unsafe_allow_html=True)
st.markdown('<div class="title">設定（マスタデータ更新）</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">商品マスタCSVファイルをアップロードして、マスタデータを更新できます。</div>', unsafe_allow_html=True)

# 商品マスタCSVのパス（pagesフォルダから見て一つ上の階層にあることを想定）
script_dir = os.path.dirname(os.path.abspath(__file__))
master_csv_path = os.path.join(script_dir, "..", "商品マスタ一覧.csv")

st.info("ここにアップロードされたCSVファイルは、現在の商品マスタデータとして保存されます。")

uploaded_master_csv = st.file_uploader(
    "新しい商品マスタCSVファイルをアップロードしてください",
    type="csv",
    help="新しいマスタデータとして使用するCSVファイルをアップロードしてください。既存の「商品マスタ一覧.csv」が上書きされます。"
)

if uploaded_master_csv is not None:
    try:
        # アップロードされたCSVをDataFrameとして読み込む
        # エンコーディングを自動判別するため、io.StringIOを使用
        csv_string = io.StringIO(uploaded_master_csv.getvalue().decode('utf-8'))
        uploaded_df = pd.read_csv(csv_string)

        # 必要な列が存在するか確認
        if '商品予定名' not in uploaded_df.columns:
            st.error("エラー: アップロードされたCSVには「商品予定名」列が必要です。")
        else:
            # CSVをバイトとして読み込み、ファイルに書き込む
            # DataFrameを直接CSVとして保存する場合、エンコーディングを指定
            uploaded_df.to_csv(master_csv_path, index=False, encoding='utf-8-sig') # UTF-8 BOM付きで保存
            
            st.success("商品マスタCSVが正常に更新されました！")
            st.write("更新されたマスタデータのプレビュー:")
            st.dataframe(uploaded_df.head()) # アップロードされたデータの先頭を表示
            st.info(f"更新されたファイル: **{os.path.basename(master_csv_path)}**")

            # Streamlit Cloudなどの環境でファイルが永続化されない場合の注意喚起
            # 環境変数 'DYNO' はHerokuでよく使われるが、Streamlit Cloudでは異なる可能性がある
            # より一般的な方法として、クラウド環境でのファイル書き込みは注意が必要であることを伝える
            st.warning("⚠️ **注意**: Streamlit Cloudのようなデプロイ環境では、この方法でアップロードされたファイルへの変更は、アプリの再起動時に失われる可能性があります。永続的なデータ更新が必要な場合は、Google DriveやS3などの外部ストレージとの連携をご検討ください。")

    except pd.errors.EmptyDataError:
        st.error("エラー: アップロードされたCSVファイルは空です。有効なデータを含むファイルをアップロードしてください。")
    except UnicodeDecodeError:
        st.error("エラー: CSVファイルのエンコーディングを判別できませんでした。UTF-8またはShift-JISで保存されたCSVファイルを使用してください。")
    except Exception as e:
        st.error(f"CSVファイルの処理中にエラーが発生しました: {e}")
        st.exception(e)

st.markdown('<div class="separator"></div>', unsafe_allow_html=True)

st.markdown("### 現在の商品マスタデータ")
if os.path.exists(master_csv_path):
    try:
        # 現在の商品マスタCSVを読み込み、表示
        # ここでもエンコーディングを複数試すロジックを再利用しても良い
        current_master_df = None
        encodings = ['utf-8', 'shift_jis', 'cp932', 'euc-jp', 'iso-2022-jp', 'utf-8-sig'] # utf-8-sigを追加
        for encoding in encodings:
            try:
                temp_df = pd.read_csv(master_csv_path, encoding=encoding)
                if not temp_df.empty:
                    current_master_df = temp_df
                    break
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
            except Exception as e:
                continue

        if current_master_df is not None and not current_master_df.empty:
            st.dataframe(current_master_df)
        else:
            st.info("現在、商品マスタデータは空、または読み込みに失敗しました。新しいCSVをアップロードしてください。")
    except Exception as e:
        st.error(f"現在の商品マスタCSVの読み込み中にエラーが発生しました: {e}")
        st.exception(e)
else:
    st.warning("商品マスタCSVファイルが見つかりません。")

st.markdown('</div>', unsafe_allow_html=True)
