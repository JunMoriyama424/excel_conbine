import pandas as pd
from openpyxl import Workbook

# 2つのExcelファイルを読み込む
excel1_path = "Excel1.xlsx"
excel2_path = "Excel2.xlsx"

# シート1をDataFrameとして読み込む
sheet1_excel1 = pd.read_excel(excel1_path, sheet_name="シート1", header=None)
sheet1_excel2 = pd.read_excel(excel2_path, sheet_name="シート1", header=None)

# シート1のデータを上下に結合（空白行を無視）
combined_sheet1 = pd.concat([sheet1_excel1, sheet1_excel2], ignore_index=True)

# 新しいExcelブックを作成
output_path = "Combined.xlsx"
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    # 結合したデータを新しいシートに書き込む
    combined_sheet1.to_excel(writer, sheet_name="シート1", index=False, header=False)

print(f"2つのExcelファイルのシート1を結合した新しいExcelファイルを作成しました: {output_path}")