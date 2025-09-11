# streamlit_app.py の修正部分

# 既存のインポートに追加
from pdf_utils import (
    safe_write_df, pdf_to_excel_data_for_paste_sheet, extract_table_from_pdf_for_bento,
    find_correct_anchor_for_bento, extract_bento_range_for_bento, match_bento_names,
    extract_detailed_client_info_from_pdf, export_detailed_client_data_to_dataframe,
    # 新しく追加する関数
    improved_pdf_to_excel_data_for_paste_sheet, debug_pdf_content
)

# PDFアップロード後の処理部分を以下のように修正：

if uploaded_pdf is not None:
    template_path = "template.xlsm"
    nouhinsyo_path = "nouhinsyo.xlsx"
    if not os.path.exists(template_path) or not os.path.exists(nouhinsyo_path):
        st.error(f"必要なテンプレートファイルが見つかりません：'{template_path}' または '{nouhinsyo_path}'")
        st.stop()
    
    template_wb = load_workbook(template_path, keep_vba=True)
    nouhinsyo_wb = load_workbook(nouhinsyo_path)
    pdf_bytes_io = io.BytesIO(uploaded_pdf.getvalue())
    
    # デバッグ情報の表示（オプション）
    if st.checkbox("PDFの詳細情報を表示"):
        with st.expander("PDFデバッグ情報"):
            debug_info = debug_pdf_content(io.BytesIO(pdf_bytes_io.getvalue()))
            st.json(debug_info)
    
    df_paste_sheet, df_bento_sheet, df_client_sheet = None, None, None
    with st.spinner("PDFからデータを抽出中..."):
        try:
            # 改善された抽出関数を使用
            df_paste_sheet = improved_pdf_to_excel_data_for_paste_sheet(io.BytesIO(pdf_bytes_io.getvalue()))
            
            # 従来の方法もフォールバックとして試行
            if df_paste_sheet is None or df_paste_sheet.empty:
                st.warning("改善された抽出方法が失敗しました。従来の方法を試行します...")
                df_paste_sheet = pdf_to_excel_data_for_paste_sheet(io.BytesIO(pdf_bytes_io.getvalue()))
                
        except Exception as e:
            st.error(f"PDFからの貼り付け用データ抽出中にエラーが発生しました: {e}")
            df_paste_sheet = None

        # 抽出結果の確認
        if df_paste_sheet is not None and not df_paste_sheet.empty:
            st.success(f"✅ データを抽出しました（{len(df_paste_sheet)}行）")
            
            # データのプレビュー表示
            if st.checkbox("抽出データをプレビュー"):
                st.dataframe(df_paste_sheet.head(10))
        else:
            st.warning("⚠️ 抽出されたデータがありません")

        # 以下、既存の弁当データとクライアント情報の処理は同様...
        if df_paste_sheet is not None:
            # 注文弁当データの抽出
            try:
                tables = extract_table_from_pdf_for_bento(io.BytesIO(pdf_bytes_io.getvalue()))
                if tables:
                    main_table = max(tables, key=len)
                    anchor_col = find_correct_anchor_for_bento(main_table)
                    if anchor_col != -1:
                        bento_list = extract_bento_range_for_bento(main_table, anchor_col)
                        if bento_list:
                            matched_list = match_bento_names(bento_list, st.session_state.master_df)
                            output_data = []
                            master_df = st.session_state.master_df

                            has_enough_columns = len(master_df.columns) > 17
                            col_p_name = master_df.columns[15] if has_enough_columns else '追加データC'
                            col_r_name = master_df.columns[17] if has_enough_columns else '追加データD'

                            for item in matched_list:
                                bento_name = ""
                                bento_iri = ""
                                match = re.search(r' \(入数: (.+?)\)$', item)
                                if match:
                                    bento_name = item[:match.start()]
                                    bento_iri = match.group(1)
                                else:
                                    bento_name = item.replace(" (未マッチ)", "")

                                val_p = ""
                                val_r = ""
                                
                                if '商品予定名' in master_df.columns:
                                    matched_row = master_df[master_df['商品予定名'] == bento_name]
                                    if not matched_row.empty and has_enough_columns:
                                        val_p = matched_row.iloc[0, 15]
                                        val_r = matched_row.iloc[0, 17]
                                
                                output_data.append([bento_name, bento_iri, val_p, val_r])
                            
                            df_bento_sheet = pd.DataFrame(output_data, columns=['商品予定名', 'パン箱入数', col_p_name, col_r_name])
            except Exception as e:
                st.error(f"注文弁当データ処理中にエラーが発生しました: {e}")

            # クライアント情報の抽出
            try:
                client_data = extract_detailed_client_info_from_pdf(io.BytesIO(pdf_bytes_io.getvalue()))
                if client_data:
                    df_client_sheet = export_detailed_client_data_to_dataframe(client_data)
                    st.success(f"クライアント情報 {len(client_data)} 件を抽出しました")
            except Exception as e:
                st.error(f"クライアント情報抽出中にエラーが発生しました: {e}")
    
    # 以下、Excelファイル生成処理は既存のまま...
