import pandas as pd
import argparse
from openpyxl import Workbook

def combine_excel_sheets(file_paths, sheet_name, output_path):
    # 各Excelファイルの指定されたシートを読み込み、結合
    combined_data = pd.DataFrame()
    for file_path in file_paths:
        sheet_data = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        combined_data = pd.concat([combined_data, sheet_data], ignore_index=True)

    # 結合したデータを新しいExcelファイルに保存
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        combined_data.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    print(f"結合したデータを新しいExcelファイルに保存しました: {output_path}")

def main():
    # コマンドライン引数の設定
    parser = argparse.ArgumentParser(description="複数のExcelファイルの指定されたシートを結合します。")
    parser.add_argument(
        "files", nargs="+", help="結合するExcelファイルのパスを指定してください（複数可）。"
    )
    parser.add_argument(
        "--sheet", required=True, help="結合するシート名を指定してください。"
    )
    parser.add_argument(
        "--output", default="Combined.xlsx", help="出力するExcelファイルのパスを指定してください（デフォルト: Combined.xlsx）。"
    )
    args = parser.parse_args()

    # 結合処理を実行
    combine_excel_sheets(args.files, args.sheet, args.output)

if __name__ == "__main__":
    main()