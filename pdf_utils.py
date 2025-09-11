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

    # --- 文字化けした列名に対応（実際の列位置で判定） ---
    if len(master_df.columns) < 18:
        return [[name, "", "マスタ列不足: 必要な列数が不足", ""] for name in pdf_bento_list]
    
    # 列のインデックスで直接アクセス（文字化け対応）
    # 想定：C列=商品名(2), E列=パン箱入数(4), P列=クラス分け名称4(15), R列=クラス分け名称5(17)
    product_name_col = master_df.columns[2]    # C列（商品名）
    pan_box_col = master_df.columns[4]         # E列（パン箱入数）
    class4_col = master_df.columns[15]         # P列（クラス分け名称4）
    class5_col = master_df.columns[17]         # R列（クラス分け名称5）
    
    # データを取得
    master_data = []
    for _, row in master_df.iterrows():
        master_data.append({
            'name': str(row[product_name_col]).strip(),
            'pan_box': str(row[pan_box_col]).strip(),
            'class4': str(row[class4_col]).strip(),
            'class5': str(row[class5_col]).strip()
        })
    
    matched_results = []
    
    for pdf_name in pdf_bento_list:
        pdf_name_stripped = pdf_name.strip()
        norm_pdf = unicodedata.normalize('NFKC', pdf_name_stripped).replace(" ", "")
        
        result_data = [pdf_name_stripped, "", "", ""]  # デフォルト値
        
        # --- マッチングロジック（完全一致を優先） ---
        best_match = None
        for item in master_data:
            norm_master = unicodedata.normalize('NFKC', item['name']).replace(" ", "")
            
            if norm_master == norm_pdf:
                best_match = [item['name'], item['pan_box'], item['class4'], item['class5']]
                break
        
        # 完全一致がなければ部分一致（含まれるか）を試す
        if not best_match:
            candidates = []
            for item in master_data:
                norm_master = unicodedata.normalize('NFKC', item['name']).replace(" ", "")
                if norm_master and norm_master in norm_pdf:  # マスタ名がPDF名に含まれるか
                    candidates.append([item['name'], item['pan_box'], item['class4'], item['class5']])
            
            if candidates:
                # 候補の中から最も長いものを採用
                best_match = max(candidates, key=lambda x: len(x[0]))

        if best_match:
            result_data = best_match  # [名前, 入数, 名称4, 名称5]
        
        matched_results.append(result_data)
        
    return matched_results
