import typer
from logging import getLogger
from util.logger import setup_root_logger
import openpyxl
import re
import os
import win32com.client
from datetime import datetime, timedelta

logger = getLogger()
app = typer.Typer()

def generate_expected_sheet_names(start_yyyymm: str, end_yyyymm: str):
    """指定された年月の範囲で、各週の月曜始まり・日曜終わりのシート名リストを生成"""
    expected_sheets = []
    
    start_date = datetime.strptime(start_yyyymm, "%Y%m").replace(day=1)
    end_date = datetime.strptime(end_yyyymm, "%Y%m").replace(day=1)
    
    next_month = end_date.month % 12 + 1
    next_month_year = end_date.year + (1 if next_month == 1 else 0)
    last_day = datetime(next_month_year, next_month, 1) - timedelta(days=1)
    
    current_date = start_date

    while current_date <= last_day:
        if current_date.weekday() != 0:
            current_date += timedelta(days=(7 - current_date.weekday()))
        
        if current_date > last_day:
            break
        
        sunday = current_date + timedelta(days=6)
        expected_sheets.append(f"{current_date.strftime('%Y%m%d')}_{sunday.strftime('%Y%m%d')}")
        
        current_date += timedelta(days=7)

    return expected_sheets

def check_sheet_dates(ws, sheet_name):
    """指定シートの H10 〜 H16 に正しい日付・数式が入っているか確認"""
    match = re.match(r"^(\d{8})_(\d{8})$", sheet_name)
    if not match:
        logger.warning(f"シート名が不正です: {sheet_name}")
        return False

    start_date = datetime.strptime(match.group(1), "%Y%m%d")

    # H10 の日付チェック
    h10_value = ws["H10"].value or ws["H10"].internal_value
    if isinstance(h10_value, datetime):
        h10_date = h10_value
    else:
        try:
            h10_date = datetime.strptime(str(h10_value), "%Y/%m/%d")
        except ValueError:
            logger.warning(f"シート {sheet_name} の H10 の値が不正: {h10_value}")
            return False

    if h10_date != start_date:
        logger.warning(f"シート {sheet_name} の H10 の値が正しくありません: {h10_date} (期待値: {start_date.strftime('%Y/%m/%d')})")
        return False

    incorrect_cells = []

    # H11～H16 の数式チェック
    for i in range(1, 7):
        cell_ref = f"H{10 + i}"
        cell_value = ws[cell_ref].value
        expected_formula = f"=H{9 + i}+1"

        if not isinstance(cell_value, str) or not cell_value.startswith("="):
            incorrect_cells.append(f"{cell_ref}: 数式が設定されていません (期待値: {expected_formula})")
        elif cell_value != expected_formula:
            incorrect_cells.append(f"{cell_ref}: {cell_value} (期待値: {expected_formula})")

    # A列の数式チェック (H10 ～ H16 を参照)
    for i, h_row in enumerate(range(10, 17)):
        a_row = 10 + (i * 6)
        h_ref = f"H{h_row}"

        expected_month_formula = f"=MONTH({h_ref})"
        expected_day_formula = f"=DAY({h_ref})"
        expected_text_formula = f'="("&TEXT({h_ref}, "aaa")&")"'

        if ws[f"A{a_row}"].value != expected_month_formula:
            incorrect_cells.append(f"A{a_row}: 数式が正しくありません (期待値: {expected_month_formula})")

        if ws[f"A{a_row + 1}"].value != expected_day_formula:
            incorrect_cells.append(f"A{a_row + 1}: 数式が正しくありません (期待値: {expected_day_formula})")

        if ws[f"A{a_row + 2}"].value != expected_text_formula:
            incorrect_cells.append(f"A{a_row + 2}: 数式が正しくありません (期待値: {expected_text_formula})")

    # B列の固定値チェック (指定セル範囲で完全一致)
    expected_b_values = {
        10: "月", 11: "日",
        16: "月", 17: "日",
        22: "月", 23: "日",
        28: "月", 29: "日",
        34: "月", 35: "日",
        40: "月", 41: "日",
        46: "月", 47: "日",
    }

    for row, expected_value in expected_b_values.items():
        cell_ref = f"B{row}"
        actual_value = ws[cell_ref].value

        if actual_value != expected_value:
            incorrect_cells.append(f"{cell_ref}: 値が正しくありません (期待値: '{expected_value}', 実際: '{actual_value}')")

    if incorrect_cells:
        logger.warning(f"シート {sheet_name} のセル値が正しくありません:\n" + "\n".join(incorrect_cells))
        return False

    return True




def check_daily_report(ws, sheet_name):
    """C9, C15, C21, C27, C33 に入力があるかチェック"""
    required_cells = ["C9", "C15", "C21", "C27", "C33"]
    empty_cells = [cell for cell in required_cells if not ws[cell].value]

    if empty_cells:
        logger.warning(f"シート {sheet_name} の日報が入力されていません: {', '.join(empty_cells)}")
        return False

    return True

@app.command("check")
def sheet_name_check(start_yyyymm: str, end_yyyymm: str):
    wb = openpyxl.load_workbook("input\日報_山田翔太 (7).xlsx", data_only=False)  # 数式を取得するため data_only=False
    existing_sheets = {ws.title for ws in wb.worksheets}

    expected_sheets = generate_expected_sheet_names(start_yyyymm, end_yyyymm)

    missing_sheets = [sheet for sheet in expected_sheets if sheet not in existing_sheets]
    extra_sheets = [sheet for sheet in existing_sheets if not re.match(r'^\d{8}_\d{8}$', sheet)]

    if missing_sheets:
        logger.warning(f"不足しているシートがあります: {missing_sheets}")
    else:
        logger.info("必要なシートはすべて存在しています。")

    if extra_sheets:
        logger.warning(f"適切ではないシート名が検出されました: {extra_sheets}")

    for sheet_name in expected_sheets:
        if sheet_name in existing_sheets:
            ws = wb[sheet_name]
            check_sheet_dates(ws, sheet_name)
            check_daily_report(ws, sheet_name)

typer_app = typer.Typer()
logger = getLogger()

@app.command("cut")
def cut_out_sheet(start_yyyymm: str, end_yyyymm: str):
    """
    指定した年月範囲に含まれるシートのみを抽出し、新しいExcelファイルとして出力する。
    """
    input_path = os.path.abspath("target\\日報_山田翔太 (7).xlsx")
    output_dir = os.path.abspath("output")
    output_filename = os.path.join(output_dir, f"日報_抽出_{start_yyyymm}_{end_yyyymm}.xlsx")
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        # Excelアプリケーションの起動
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Excelを非表示で実行

        # 元のExcelファイルを開く
        wb = excel.Workbooks.Open(input_path)
        
        # 新規Excelファイルを作成
        new_wb = excel.Workbooks.Add()

        # 期待されるシート名を取得（順番通り）
        expected_sheets = generate_expected_sheet_names(start_yyyymm, end_yyyymm)

        # シートを順番通りにコピー
        copied_any = False
        for sheet_name in expected_sheets:
            if sheet_name in [sheet.Name for sheet in wb.Sheets]:
                wb.Sheets(sheet_name).Copy(Before=new_wb.Sheets(1))
                copied_any = True

        if copied_any:
            # 最初に自動作成された空のシートがある場合は削除
            if new_wb.Sheets.Count > 1:
                new_wb.Sheets(new_wb.Sheets.Count).Delete()

            # 新規Excelファイルを保存
            new_wb.SaveAs(output_filename)
            logger.info(f"抽出したシートを {output_filename} に出力しました。")
        else:
            logger.warning("指定範囲に該当するシートがありませんでした。")

        # ファイルを閉じる
        new_wb.Close(SaveChanges=False)
        wb.Close(SaveChanges=False)

    except Exception as e:
        logger.error(f"エラーが発生しました: {e}")
    
    finally:
        excel.Quit()  # Excelを終了

if __name__ == "__main__":
    setup_root_logger(verbose=True)
    app()
