import pandas as pd
from openpyxl import Workbook

# 2つのExcelファイルを読み込む
excel1_path = "Excel1.xlsx"
excel2_path = "Excel2.xlsx"

# シート1をDataFrameとして読み込む
sheet1_excel1 = pd.read_excel(excel1_path, sheet_name="シート1")
sheet1_excel2 = pd.read_excel(excel2_path, sheet_name="シート1")

# 新しいExcelブックを作成
output_path = "Combined.xlsx"
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    # シート1のデータをそれぞれ新しいシートに書き込む
    sheet1_excel1.to_excel(writer, sheet_name="Excel1_Sheet1", index=False)
    sheet1_excel2.to_excel(writer, sheet_name="Excel2_Sheet1", index=False)

print(f"2つのExcelファイルのシート1を結合した新しいExcelファイルを作成しました: {output_path}")