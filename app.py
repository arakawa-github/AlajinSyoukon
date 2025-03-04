import os
import pandas as pd
from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename
import re

app = Flask(__name__)

# アップロードフォルダ設定
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# 許可する拡張子
ALLOWED_EXTENSIONS = {"xlsx", "xls"}
def alajin2(df_a,df_b_template):
    # A列 → B列 のマッピング
    mapping = {
        "売上日": "売上日",
        "出荷日":"請求日",
        "受注NO": "伝票No",
        "得意先": "得意先コード",
        "得意先略称": "摘要名",
        "商品": "商品名",
        "売上数": "数量",
        "売上単価": "単価",
        "売上金額": "売上金額",
        "原価単価": "原単価",
        "原価金額": "原価金額",
        "相手先商品コード":"備考"
    }

    # df_b_template の列名とデータ型を保持した空の DataFrame を作成
    df_b = pd.DataFrame(columns=df_b_template.columns)

    # データの更新
    for index, row in df_a.iterrows():
        new_row = {}
        for a_col, b_col in mapping.items():
            if a_col in df_a.columns and b_col in df_b.columns:
                # 日付を YYYYMMDD 形式に変換
                if a_col in ["売上日", "出荷日"]:
                
                    date_value = pd.to_datetime(row[a_col], errors="coerce")
                    new_row[b_col] = date_value.strftime("%Y%m%d") if not pd.isna(date_value) else ""
                else:
                    new_row[b_col] = row[a_col]
             
                if a_col== "受注NO" :
                    n_values = new_row[b_col]# n_valuesは数値  小数点1桁                                         
                    new_row[b_col] = str(int(n_values) if pd.notna(n_values) else 0).zfill(4)[-4:]
                # 「得意先」の ０１０をNに置き換え 
            
                if a_col == "得意先":               
                    # もし new_row[b_col] が数値だった場合、文字列に変換
                    #str_value = str(new_row[b_col])
                    str_value = str(int(row[a_col])) if not pd.isna(row[a_col]) else ""                
                
                    # 先頭の "010" を削除し、前に "N" を追加 010が１０となっているため、
                    #変換すると、010でなく、10となっている。
                
                    if str_value=="1020161":  #小林製作所"
                        #print("index"+str(index)+"得意先"+str_value)
                        new_row[b_col] = "N20160"
                    else:                    
                        if str_value=="1005004":  #三菱ロジ"
                            #print("index"+str(index)+"得意先"+str_value)
                            new_row[b_col] = "N50040"
                        else:
                            # 先頭の "010" を削除し、前に "N" を追加 010が１０となっているため、
                            new_row[b_col] = "N" + re.sub(r"^10", "", str_value)     
                    
                # 数値を整数型に変換（売上数・売上金額・粗利益）
                if a_col in ["売上数", "売上金額", "粗利益"]:
                    new_row[b_col] = int(float(new_row[b_col])) if pd.notna(new_row[b_col]) else 0

        # 新しい行を df_b に追加
        df_b = pd.concat([df_b, pd.DataFrame([new_row])], ignore_index=True)

    zero_columns = [
        "伝区","マスター区分", "区", "入数","箱数","標準価格","同時入荷区分","売単価",
        "売価金額","計算式コード","商品項目１","商品項目２","商品項目３","売上項目１","売上項目２",
        "売上項目３","伝票消費税","データ区分", "単位区分", "決裁日","決裁手数料","手数料税率"
    ]
    for col in zero_columns:
        if col in df_b.columns:
            df_b[col] = 0

    # 「商品」列を 99 に設定
    if "商品名" in df_b.columns:
        df_b["商品"] = "99"
    
    if "担当者コード" in df_b.columns:
        df_b["担当者コード"] = "0039"  # 文字列として代入

    if "部門コード" in df_b.columns:
        df_b["部門コード"] = "005"  # 文字列として代入

    if "税率" in df_b.columns:
        df_b["税率"] = "10"  # 文字列として代入

    return df_b

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_files():
    if "file1" not in request.files or "file2" not in request.files:
        return "2つのファイルを選択してください", 400

    file1 = request.files["file1"]
    file2 = request.files["file2"]

    if file1.filename == "" or file2.filename == "":
        return "2つのファイルを選択してください", 400

    if file1 and allowed_file(file1.filename) and file2 and allowed_file(file2.filename):
        filename1 = secure_filename(file1.filename)
        filename2 = secure_filename(file2.filename)

        file_path1 = os.path.join(app.config["UPLOAD_FOLDER"], filename1)
        file_path2 = os.path.join(app.config["UPLOAD_FOLDER"], filename2)

        file1.save(file_path1)
        file2.save(file_path2)

        # 2つのファイルをDataFrameとして読み込む
        df_a = pd.read_excel(file_path1, header=6) #アラジンデータ
        df_b_template = pd.read_excel(file_path2)# format.xlsx の読み込み

        # 2つのアラジンデータ、formatデータを処理
        merged_df = alajin2(df_a,df_b_template)

        # 保存ファイル名
        output_filename = "商魂_output.csv"
        output_path = os.path.join(app.config["UPLOAD_FOLDER"], output_filename)
        # csvに保存    
        merged_df.to_csv(output_path, index=False, header=False,encoding="cp932")

        return render_template("index.html", filename=output_filename)

    return "許可されていないファイル形式です", 400

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
    #app.run(host='0.0.0.0', port=5000)
