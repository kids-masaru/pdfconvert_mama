import streamlit as st
import pandas as pd
import os

st.markdown('<div class="title">マスタデータ設定</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">更新するマスタの確認、および新しいCSVファイルのアップロードができます。</div>', unsafe_allow_html=True)

st.markdown("##### 現在の商品マスタデータ（プレビュー）")
if 'master_df' in st.session_state and not st.session_state.master_df.empty:
    st.dataframe(st.session_state.master_df.head(), use_container_width=True)
else:
    st.warning("商品マスタが読み込まれていません。")

st.markdown("##### 現在の得意先マスタデータ（プレビュー）")
if 'customer_master_df' in st.session_state and not st.session_state.customer_master_df.empty:
    st.dataframe(st.session_state.customer_master_df.head(), use_container_width=True)
else:
    st.warning("得意先マスタが読み込まれていません。")

st.markdown("---")

master_choice = st.selectbox(
    "更新するマスタを選択してください",
    ("商品マスタ", "得意先マスタ")
)

if master_choice == "商品マスタ":
    st.markdown("#### 商品マスタの更新")
    master_csv_path = "商品マスタ一覧.csv"
    uploaded_master_csv = st.file_uploader("新しい商品マスタ一覧.csvをアップロード",type="csv",help="ヘッダーには '商品予定名', 'パン箱入数', '商品名' を含めてください。",key="product_master_uploader")
    if uploaded_master_csv is not None:
        try:
            new_master_df = None
            encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
            for encoding in encodings:
                try:
                    uploaded_master_csv.seek(0)
                    temp_df = pd.read_csv(uploaded_master_csv, encoding=encoding)
                    if all(col in temp_df.columns for col in ['商品予定名', 'パン箱入数', '商品名']):
                        new_master_df = temp_df
                        st.info(f"ファイルを {encoding} で読み込みました。")
                        break
                except Exception: continue
            if new_master_df is not None:
                st.session_state.master_df = new_master_df
                new_master_df.to_csv(master_csv_path, index=False, encoding='utf-8-sig')
                st.success(f"✅ 商品マスタを更新し、'{master_csv_path}' に保存しました。")
            else:
                st.error("CSVファイルを正しく読み込めませんでした。必須列を確認してください。")
        except Exception as e:
            st.error(f"商品マスタ更新処理中にエラー: {e}")
    st.markdown("##### 現在の商品マスタデータ（全件）")
    if 'master_df' in st.session_state and not st.session_state.master_df.empty:
        st.dataframe(st.session_state.master_df, use_container_width=True)
    else:
        st.warning("商品マスタが読み込まれていません。")

elif master_choice == "得意先マスタ":
    st.markdown("#### 得意先マスタの更新")
    customer_master_csv_path = "得意先マスタ一覧.csv"
    uploaded_customer_csv = st.file_uploader("新しい得意先マスタ一覧.csvをアップロード",type="csv",help="ヘッダーには '得意先コード', '得意先名' を含めてください。",key="customer_master_uploader")
    if uploaded_customer_csv is not None:
        try:
            new_customer_df = None
            encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
            for encoding in encodings:
                try:
                    uploaded_customer_csv.seek(0)
                    temp_df = pd.read_csv(uploaded_customer_csv, encoding=encoding)
                    if all(col in temp_df.columns for col in ['得意先コード', '得意先名']):
                        new_customer_df = temp_df
                        st.info(f"ファイルを {encoding} で読み込みました。")
                        break
                except Exception: continue
            if new_customer_df is not None:
                st.session_state.customer_master_df = new_customer_df
                new_customer_df.to_csv(customer_master_csv_path, index=False, encoding='utf-8-sig')
                st.success(f"✅ 得意先マスタを更新し、'{customer_master_csv_path}' に保存しました。")
            else:
                st.error("CSVファイルを正しく読み込めませんでした。必須列を確認してください。")
        except Exception as e:
            st.error(f"得意先マスタ更新処理中にエラー: {e}")
    st.markdown("##### 現在の得意先マスタデータ（全件）")
    if 'customer_master_df' in st.session_state and not st.session_state.customer_master_df.empty:
        st.dataframe(st.session_state.customer_master_df, use_container_width=True)
    else:
        st.warning("得意先マスタが読み込まれていません。")
