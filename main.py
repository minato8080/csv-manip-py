import csv
import os
from openpyxl import Workbook
from openpyxl.comments import Comment
from datetime import datetime, timedelta
import json

def format_csv(reader):
        # CSVの1列目の日付を収集し、日付以外のデータを弾く
        dates = []
        valid_rows = []
        for row in reader:
            try:
                # 日付をdatetimeオブジェクトに変換
                date_obj = datetime.strptime(row[0], "%Y/%m/%d")
                dates.append(date_obj)
                valid_rows.append(row)
            except ValueError:
                # 日付以外のデータが入っていた場合は無視
                continue
        
        # 最小値と最大値を取得
        if not dates:
            return json.dumps({}, ensure_ascii=False, indent=4)
        
        min_date = min(dates)
        max_date = max(dates)
        
        # 日付範囲を作成
        date_range = [min_date + timedelta(days=x) for x in range((max_date - min_date).days + 1)]
        
        # 日付をキーにしてデータをまとめる
        date_dict = {date.strftime("%Y/%m/%d"): [] for date in date_range}
        for row in valid_rows:
            date_str = row[0]
            if date_str in date_dict:
                date_dict[date_str].append(row[1:])
        
        # JSON形式にパース
        return json.dumps(date_dict, ensure_ascii=False, indent=4)

def csv_to_excel_with_comments(csv_file, excel_file):
    # CSVファイルを読み込む
    with open(csv_file, mode='r', encoding='shift_jis') as f:
        # CSVリーダーを作成
        reader = csv.reader(f)
        
        # CSVデータをフォーマットしてJSON形式に変換
        json_data = json.loads(format_csv(reader))
    
    # Excelファイルを作成
    wb = Workbook()
    ws = wb.active
    
    # データをExcelに書き込み
    for col_idx, key in enumerate(json_data, start=1):
        cell = ws.cell(row=1, column=col_idx, value=datetime.strptime(key, "%Y/%m/%d").strftime("%m/%d"))
        for row_idx, row in enumerate(json_data[key], start=2):
            cell = ws.cell(row=row_idx, column=col_idx, value=f"¥{int(row[1]):,}")
            # コメントを追加
            comment_text = f"{row[0]}"
            cell.comment = Comment(comment_text, "GeneratedByScript")
    
    # 列幅を1.2cmに設定
    for col in ws.columns:
        max_length = 1.2 / 0.18  # 1.2cmをポイントに変換（1ポイント=0.18cm）
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max_length
    
    # Excelファイルを保存
    wb.save(excel_file)
    print(f"Excelファイルを作成しました: {excel_file}")

# 使用例
import tkinter as tk
from tkinter import filedialog, messagebox

def select_files_and_execute():
    root = tk.Tk()
    root.withdraw()  # メインウィンドウを表示しない

    # CSVファイルを選択
    csv_file = filedialog.askopenfilename(title="CSVファイルを選択", filetypes=[("CSV files", "*.csv")])
    if not csv_file:
        messagebox.showinfo("キャンセル", "CSVファイルの選択がキャンセルされました。")
        return

    # デフォルトで選択したCSVファイルと同じ名前のExcelファイルを設定（フルパスを除去）
    default_excel_file = os.path.basename(csv_file).rsplit('.', 1)[0] + '.xlsx'

    # Excelファイルを保存する場所を選択
    excel_file = filedialog.asksaveasfilename(title="Excelファイルを保存", initialfile=default_excel_file, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not excel_file:
        messagebox.showinfo("キャンセル", "Excelファイルの保存がキャンセルされました。")
        return

    csv_to_excel_with_comments(csv_file, excel_file)
select_files_and_execute()
