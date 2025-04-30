import openpyxl

def load_script_list(excel_path, sheet_name='Sheet1', start_row=3):
    """
    Excel の台詞リストを読み込んで
    { '番号': '台詞本文', ... } の dict を返します。
    """
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[sheet_name]
    scripts = {}
    for row in ws.iter_rows(min_row=start_row, max_col=2, values_only=True):
        number, text = row
        if number is not None and text:
            scripts[str(number)] = text
    return scripts

