import streamlit as st
import pandas as pd
import os

# --- サイドバーの表示 ---
st.sidebar.title("メニュー")
st.sidebar.page_link("streamlit_app.py", label="PDF Excel 変換", icon="📄")
st.sidebar.page_link("pages/マスタ設定.py", label="マスタ設定", icon="⚙️")

st.markdown("""
    <!-- PWA メタタグ -->
    <link rel="manifest" href="./static/manifest.json">
    <meta name="theme-color" content="#ffffff">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="default">
    <meta name="apple-mobile-web-app-title" content="PDF変換ツール">
    <link rel="apple-touch-icon" href="./static/icons/apple-touch-icon.png">
    <link rel="icon" type="image/png" sizes="192x192" href="./static/icons/android-chrome-192.png">
    <link rel="icon" type="image/png" sizes="512x512" href="./static/icons/android-chrome-512.png">
    
    <style>
        [data-testid="stSidebarNav"] {
            display: none;
        }
        .custom-title {
            font-size: 2.1rem;
            font-weight: 600;
            color: #3A322E;
            padding-bottom: 10px;
            border-bottom: 3px solid #FF9933;
            margin-bottom: 25px;
        }
        .stApp { 
            background: #fff5e6; 
        }
    </style>
""", unsafe_allow_html=True)

# --- ここからがページ本体のコンテンツ ---
st.markdown('<div class="title">マスタデータ設定</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">更新するマスタの確認、および新しいCSVファイルのアップロードができます。</div>', unsafe_allow_html=True)

st.markdown("##### 現在の商品マスタデータ（プレビュー）")
if 'master_df' in st.session_state and not st.session_state.master_df.empty:
    st.dataframe(st.session_state.master_df.head(), width='stretch')
else:
    st.warning("商品マスタが読み込まれていません。")

st.markdown("##### 現在の得意先マスタデータ（プレビュー）")
if 'customer_master_df' in st.session_state and not st.session_state.customer_master_df.empty:
    st.dataframe(st.session_state.customer_master_df.head(), width='stretch')
else:
    st.warning("得意先マスタが読み込まれていません。")

st.markdown("---")

master_choice = st.selectbox(
    "更新するマスタを選択してください",
    ("商品マスタ", "得意先マスタ")
)

def try_read_csv_filelike(filelike, required_cols):
    """
    複数エンコーディングで読み込みを試み、必須列が揃う最初のDataFrameを返す。
    成功しなければ None を返す。内部の例外は表示しない。
    """
    encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
    for enc in encodings:
        try:
            filelike.seek(0)
            df = pd.read_csv(filelike, encoding=enc)
            if all(col in df.columns for col in required_cols):
                return df, enc
        except Exception:
            continue
    # 最後にエラー許容で一度だけ試す（不正バイトを置換）
    try:
        filelike.seek(0)
        df = pd.read_csv(filelike, encoding='utf-8', engine='python', errors='replace')
        if all(col in df.columns for col in required_cols):
            return df, 'utf-8 (replace errors)'
    except Exception:
        pass
    return None, None

if master_choice == "商品マスタ":
    st.markdown("#### 商品マスタの更新")
    master_csv_path = os.path.abspath("商品マスタ一覧.csv")  # 絶対パス使用
    uploaded_master_csv = st.file_uploader(
        "新しい商品マスタ一覧.csvをアップロード",
        type="csv",
        help="ヘッダーには '商品予定名', 'パン箱入数', '商品名' を含めてください。",
        key="product_master_uploader"
    )
    if uploaded_master_csv is not None:
        try:
            required_cols = ['商品予定名', 'パン箱入数', '商品名']
            new_master_df, used_enc = try_read_csv_filelike(uploaded_master_csv, required_cols)
            if new_master_df is not None:
                st.session_state.master_df = new_master_df

                # バックアップ作成
                if os.path.exists(master_csv_path):
                    backup_path = master_csv_path.replace('.csv', '_backup.csv')
                    os.rename(master_csv_path, backup_path)
                    st.info(f"既存ファイルをバックアップしました: {backup_path}")

                # 新しいファイルを保存（utf-8-sig）
                new_master_df.to_csv(master_csv_path, index=False, encoding='utf-8-sig')

                # 保存確認
                if os.path.exists(master_csv_path):
                    try:
                        verification_df = pd.read_csv(master_csv_path, encoding='utf-8-sig')
                        if len(verification_df) == len(new_master_df):
                            st.success(f"✅ 商品マスタを更新し、'{master_csv_path}' に正常に保存しました。")
                            st.info(f"読み込みに使用したエンコーディング: {used_enc}")
                            st.info(f"更新件数: {len(new_master_df)} 件")
                        else:
                            st.warning("ファイルは保存されましたが、データ件数が一致しません。")
                    except Exception:
                        st.warning("ファイル保存の検証中に問題が発生しました（詳細はログ）。")
                else:
                    st.error("ファイルの保存に失敗しました。")
            else:
                st.error("CSVファイルを正しく読み込めませんでした。ヘッダーやファイル形式を確認してください。")
        except Exception:
            st.error("商品マスタの更新中にエラーが発生しました。管理者にお問い合わせください。")

    st.markdown("##### 現在の商品マスタデータ（全件）")
    if 'master_df' in st.session_state and not st.session_state.master_df.empty:
        st.dataframe(st.session_state.master_df, width='stretch')
    else:
        st.warning("商品マスタが読み込まれていません。")

elif master_choice == "得意先マスタ":
    st.markdown("#### 得意先マスタの更新")
    customer_master_csv_path = os.path.abspath("得意先マスタ一覧.csv")  # 絶対パス使用
    uploaded_customer_csv = st.file_uploader(
        "新しい得意先マスタ一覧.csvをアップロード",
        type="csv",
        help="ヘッダーには '得意先ＣＤ', '得意先名' を含めてください。",
        key="customer_master_uploader"
    )
    if uploaded_customer_csv is not None:
        try:
            required_cols = ['得意先ＣＤ', '得意先名']
            new_customer_df, used_enc = try_read_csv_filelike(uploaded_customer_csv, required_cols)
            if new_customer_df is not None:
                st.session_state.customer_master_df = new_customer_df

                # バックアップ作成
                if os.path.exists(customer_master_csv_path):
                    backup_path = customer_master_csv_path.replace('.csv', '_backup.csv')
                    os.rename(customer_master_csv_path, backup_path)
                    st.info(f"既存ファイルをバックアップしました: {backup_path}")

                # 新しいファイルを保存（utf-8-sig）
                new_customer_df.to_csv(customer_master_csv_path, index=False, encoding='utf-8-sig')

                # 保存確認
                if os.path.exists(customer_master_csv_path):
                    try:
                        verification_df = pd.read_csv(customer_master_csv_path, encoding='utf-8-sig')
                        if len(verification_df) == len(new_customer_df):
                            st.success(f"✅ 得意先マスタを更新し、'{customer_master_csv_path}' に正常に保存しました。")
                            st.info(f"読み込みに使用したエンコーディング: {used_enc}")
                            st.info(f"更新件数: {len(new_customer_df)} 件")
                        else:
                            st.warning("ファイルは保存されましたが、データ件数が一致しません。")
                    except Exception:
                        st.warning("ファイル保存の検証中に問題が発生しました（詳細はログ）。")
                else:
                    st.error("ファイルの保存に失敗しました。")
            else:
                st.error("CSVファイルを正しく読み込めませんでした。ヘッダーやファイル形式を確認してください。")
        except Exception:
            st.error("得意先マスタの更新中にエラーが発生しました。管理者にお問い合わせください。")

    st.markdown("##### 現在の得意先マスタデータ（全件）")
    if 'customer_master_df' in st.session_state and not st.session_state.customer_master_df.empty:
        st.dataframe(st.session_state.customer_master_df, width='stretch')
    else:
        st.warning("得意先マスタが読み込まれていません。")
