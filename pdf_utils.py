def match_bento_data(pdf_bento_list: List[str], master_df: pd.DataFrame) -> List[List[str]]:
    """
    PDFから抽出した弁当名リストを商品マスタと照合し、
    [弁当名, パン箱入数, クラス分け名称4, クラス分け名称5] のリストを返す。
    CSVヘッダーのスペース問題と文字化け問題もここで吸収する。
    """
    if master_df is None or master_df.empty:
        return [[name, "", "", ""] for name in pdf_bento_list]

    # --- CSVヘッダーのスペース問題をここで吸収 ---
    master_df.columns = master_df.columns.str.strip()
    
    # --- デバッグ：実際の列構造を確認 ---
    print(f"列数: {len(master_df.columns)}")
    for i, col in enumerate(master_df.columns):
        print(f"列{i}: '{col}' - サンプル: '{master_df.iloc[0, i] if not master_df.empty else 'なし'}'")
    
    if len(master_df.columns) < 18:
        return [[name, "", "マスタ列不足: 必要な列数が不足", ""] for name in pdf_bento_list]
    
    matched_results = []
    
    for pdf_name in pdf_bento_list:
        pdf_name_stripped = pdf_name.strip()
        norm_pdf = unicodedata.normalize('NFKC', pdf_name_stripped).replace(" ", "")
        
        result_data = [pdf_name_stripped, "", "", ""]  # デフォルト値
        
        # 全ての行を検索してマッチするものを探す
        best_match = None
        for idx, row in master_df.iterrows():
            # 各列で商品名を検索
            for col_idx in range(min(5, len(master_df.columns))):  # 最初の5列で検索
                cell_value = str(row.iloc[col_idx]).strip()
                if cell_value and cell_value != 'nan':
                    norm_master = unicodedata.normalize('NFKC', cell_value).replace(" ", "")
                    
                    # 完全一致チェック
                    if norm_master == norm_pdf:
                        # マッチした場合、その行の各列の値を取得
                        pan_box = str(row.iloc[4]) if len(row) > 4 else ""  # E列想定
                        class4 = str(row.iloc[15]) if len(row) > 15 else ""  # P列想定  
                        class5 = str(row.iloc[17]) if len(row) > 17 else ""  # R列想定
                        
                        best_match = [cell_value, pan_box, class4, class5]
                        print(f"完全一致: {pdf_name} -> {cell_value} (行{idx}, 列{col_idx})")
                        print(f"  パン箱入数(列4): {pan_box}")
                        print(f"  クラス分け名称4(列15): {class4}")
                        print(f"  クラス分け名称5(列17): {class5}")
                        break
                        
            if best_match:
                break
        
        # 完全一致がなければ部分一致を試す
        if not best_match:
            for idx, row in master_df.iterrows():
                for col_idx in range(min(5, len(master_df.columns))):
                    cell_value = str(row.iloc[col_idx]).strip()
                    if cell_value and cell_value != 'nan':
                        norm_master = unicodedata.normalize('NFKC', cell_value).replace(" ", "")
                        
                        # 部分一致チェック (マスタ名がPDF名に含まれる)
                        if norm_master and norm_master in norm_pdf:
                            pan_box = str(row.iloc[4]) if len(row) > 4 else ""
                            class4 = str(row.iloc[15]) if len(row) > 15 else ""
                            class5 = str(row.iloc[17]) if len(row) > 17 else ""
                            
                            best_match = [cell_value, pan_box, class4, class5]
                            print(f"部分一致: {pdf_name} -> {cell_value} (行{idx}, 列{col_idx})")
                            break
                            
                if best_match:
                    break

        if best_match:
            result_data = best_match
        
        matched_results.append(result_data)
        
    return matched_results
