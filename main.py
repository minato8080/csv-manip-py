import csv
import os
from openpyxl import Workbook
from openpyxl.comments import Comment
from datetime import datetime, timedelta
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

global default_excel_file
global csv_file
global file_label
global tree


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
    date_range = [
        min_date + timedelta(days=x) for x in range((max_date - min_date).days + 1)
    ]

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
    with open(csv_file, mode="r", encoding="shift_jis") as f:
        # CSVリーダーを作成
        reader = csv.reader(f)

        # CSVデータをフォーマットしてJSON形式に変換
        json_data = json.loads(format_csv(reader))

    # Excelファイルを作成
    wb = Workbook()
    ws = wb.active

    # データをExcelに書き込み
    for col_idx, key in enumerate(json_data, start=1):
        cell = ws.cell(
            row=1,
            column=col_idx,
            value=datetime.strptime(key, "%Y/%m/%d").strftime("%m/%d"),
        )
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
    messagebox.showinfo("完了", f"Excelファイルを作成しました: {excel_file}")


def csv_import():
    global default_excel_file
    global csv_file
    global tree
    csv_file = filedialog.askopenfilename(
        title="CSVファイルを選択", filetypes=[("CSV files", "*.csv")]
    )
    if not csv_file:
        messagebox.showinfo("キャンセル", "CSVファイルの選択がキャンセルされました。")
        return
    # デフォルトで選択したCSVファイルと同じ名前のExcelファイルを設定（フルパスを除去）
    default_excel_file = os.path.basename(csv_file).rsplit(".", 1)[0] + ".xlsx"
    file_label.config(text=f"選択ファイル: {csv_file}")

    for item in tree.get_children():
        tree.delete(item)  # 既存の内容をクリア

    with open(csv_file, mode="r", encoding="shift_jis") as f:
        reader = csv.reader(f)
        max_columns = 0
        # Treeviewに挿入
        for row in reader:
            if max_columns < len(row):
                max_columns = len(row)
            tree.insert("", "end", values=row)

        # ヘッダーを設定
        for idx in range(max_columns):
            tree.heading(f"#{idx}", text=idx, anchor="w")
            tree.column(f"#{idx}", anchor="w")


def csv_export():
    global default_excel_file
    global csv_file
    excel_file = filedialog.asksaveasfilename(
        title="Excelファイルを保存",
        initialfile=default_excel_file,
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
    )
    if not excel_file:
        messagebox.showinfo("キャンセル", "Excelファイルの保存がキャンセルされました。")
        return
    csv_to_excel_with_comments(csv_file, excel_file)


def main():
    global file_label
    global csv_file
    global tree
    root = tk.Tk()
    root.title("CSV to Excel Converter")

    # フレームを作成
    frame = tk.Frame(root, width=400, height=300)
    frame.grid(padx=10, pady=10)

    # CSVファイル名を表示するラベルを作成
    file_label = tk.Label(frame, text="選択されたCSVファイル: なし", width=50)
    file_label.grid(row=0, column=2, padx=5, pady=5)

    # CSVインポートボタンを作成
    import_button = tk.Button(frame, text="CSV Import", command=csv_import)
    import_button.grid(row=0, column=0, padx=1, pady=2)

    # CSVエクスポートボタンを作成
    export_button = tk.Button(frame, text="CSV Export", command=csv_export)
    export_button.grid(row=0, column=1, padx=1, pady=2)

    # フレームを作成
    frame2 = tk.Frame(root, width=840, height=200)
    frame2.grid(padx=10, pady=10)
    # 1列目を可変サイズとする
    frame2.columnconfigure(0, weight=1)
    # 1行目を可変サイズとする
    frame2.rowconfigure(0, weight=1)
    # 内部のサイズに合わせたフレームサイズとしない
    frame2.grid_propagate(False)
    # グリッドビューを初期化
    tree = ttk.Treeview(frame2, columns=("", "", "", ""), show="headings", height=15)
    tree.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

    v_scrollbar = ttk.Scrollbar(frame2, orient=tk.VERTICAL, command=tree.yview)
    h_scrollbar = ttk.Scrollbar(frame2, orient=tk.HORIZONTAL, command=tree.xview)
    h_scrollbar.grid(row=1, column=0, sticky=tk.EW)
    v_scrollbar.grid(row=0, column=1, sticky=tk.NS)
    tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

    root.mainloop()


main()
