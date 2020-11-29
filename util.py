# モジュール読み込み
import win32com.client
import os
from pathlib import Path
from pdf2image import convert_from_path
from pdf2image.exceptions import PDFPageCountError


# ExcelからPDFに変換し、保存する関数。
def excel2image(before_excel, after_excel, cur_path):
    # Excelを起動
    excel = win32com.client.Dispatch("Excel.Application")
    file_before = excel.Workbooks.Open(before_excel)
    file_before.WorkSheets(1).Select()
    file_before.ActiveSheet.ExportAsFixedFormat(0, cur_path+'/before.pdf')
    excel.Quit()

    excel = win32com.client.Dispatch("Excel.Application")
    file_after = excel.Workbooks.Open(after_excel)
    file_after.WorkSheets(1).Select()
    file_after.ActiveSheet.ExportAsFixedFormat(0, cur_path+'/after.pdf')
    excel.Quit()

    pdf2image_func('./before.pdf', './after.pdf')

    return 0


# PDFを画像に変換する関数。
def pdf2image_func(before_pdf, after_pdf):
    # poppler/binを環境変数PATHに追加する
    poppler_dir = Path(__file__).parent.absolute() / "poppler/bin"
    os.environ["PATH"] += os.pathsep + str(poppler_dir)

    # PDFファイルのパス
    pdf_path_before = Path(before_pdf)
    pdf_path_after = Path(after_pdf)

    # PDF -> Image に変換（150dpi）
    pages_before = convert_from_path(str(pdf_path_before), 150)
    pages_after = convert_from_path(str(pdf_path_after), 150)

    # 画像ファイルを１ページずつ保存
    os.makedirs('./image_file', exist_ok=True)
    image_dir = Path("./image_file")
    for i, page in enumerate(pages_before):
        file_name_before = 'before' + "_{:02d}".format(i + 1) + ".png"
        image_path_before = image_dir / file_name_before
        # JPEGで保存
        page.save(str(image_path_before), "PNG")

    for i, page in enumerate(pages_after):
        file_name_after = 'after' + "_{:02d}".format(i + 1) + ".png"
        image_path_after = image_dir / file_name_after
        # JPEGで保存
        page.save(str(image_path_after), "PNG")

    return 0
