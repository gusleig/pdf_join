import glob
import sys
import os
import argparse
import win32com.client
from PyPDF2 import PdfFileMerger, PdfFileReader

ROOT_DIR = os.path.abspath(os.curdir)
# Path of the pdf
PDF_FOLDER = ROOT_DIR + r"\PDF"


def xlsx_pdf(xlsx, pdf_output):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        print('Start conversion to PDF')
        # Open
        wb = excel.Workbooks.Open(xlsx)
        # Use first sheet
        ws_index_list = [1]
        wb.WorkSheets(ws_index_list).Select()
        # Save
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_output)
    except Exception as e:
        print('failed.')
    else:
        print('Succeeded.')
    finally:
        wb.Close()
        excel.Quit()


def pdf_join(pdf_list, output):

    merger = PdfFileMerger()

    for pdf in pdf_list:
        merger.append(pdf)

    merger.write(output)
    merger.close()


if __name__ == '__main__':

    parser = argparse.ArgumentParser()

    parser.add_argument('-f', '--fpath', help='files path', default="")

    args = parser.parse_args()

    fpath = args.fpath

    if not fpath:
        fpath = PDF_FOLDER

    if not os.path.isdir(fpath):
        print(f"Pasta origem inválida: {fpath}")
        sys.exit(-1)
    else:
        xlsx_files = glob.glob(fpath + "\\*.xlsx")
        pdf_files = glob.glob(fpath + "\\*.pdf")

    if xlsx_files:
        xls_file_name = os.path.basename(xlsx_files[0])
        xlsx_pdf(xlsx_files[0], os.path.join(fpath, "1.output.pdf"))
    else:
        print(f"No xlsx files in source folder: {fpath}")

    print(f"Using folder: {fpath}")

    pdf_files = glob.glob(fpath + "\\*.pdf")

    pdf_join(pdf_files, os.path.join(fpath, "1.output-final.pdf"))


