import typer
from logging import getLogger
from util.logger import setup_root_logger
import openpyxl
import re
from datetime import datetime

logger = getLogger()
app= typer.Typer()

@app.command()
def test():
    logger.info("test")

    
@app.command()
def sheet_name_check():
    wb = openpyxl.load_workbook("target\日報_山田翔太 (7).xlsx")
    ws_list = wb.worksheets
    # 正規表現パターンを定義
    pattern = r'^\d{8}_\d{8}$'  # YYYYMMDD_YYYYMMDD形式
    error_list= []
    
    # シート名から日付を取り出す
    for sheet_name in ws_list:
        if re.match(pattern, sheet_name.title):
            pass
        else:
            error_list.append(sheet_name.title)
    for i in error_list:
        logger.info(f"適切ではないシート名が検出されました: {i}")



if __name__ == "__main__":
    setup_root_logger(verbose=True)
    app()
