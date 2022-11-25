import glob
import sys
import os
import argparse
import re
import win32com.client
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfWriter, PdfReader, PdfFileWriter
# from pdf_compressor import CompressPDF
import subprocess
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


def pdf_compress(pdf_name):

    reader = PdfReader(pdf_name)

    pdf_writer = PdfFileWriter()

    for page in reader.pages:

        page.compress_content_streams()
        pdf_writer.add_page(page)

    with open(pdf_name.replace(".pdf", "_compressed.pdf"), "wb") as f:
        pdf_writer.write(f)


def compress_pdf(path):

    # compress = 4 # screen
    compress = 2  # printer

    p = CompressPDF(compress, show_info=True)

    pdf_files = glob.glob(path + "*.pdf")

    for pdf_file in pdf_files:
        new_file = pdf_file.replace(".pdf", "_compressed.pdf")

        pdf_file_name = pdf_file[pdf_file.rfind("\\") + 1:]

        # new_file = os.path.join(compress_folder, filename)
        try:
            if p.compress(pdf_file, new_file):
                print("{} done!".format(pdf_file_name))
            else:
                print("{} gave an error!".format(pdf_file))
        except Exception as e:
            print(str(e))


def pdf_join(pdf_list, output):

    merger = PdfFileMerger()

    for pdf in pdf_list:
        merger.append(pdf)

    merger.write(output)
    merger.close()


def delete_file(xls_new_pdf_name):

    ## If file exists, delete it ##
    if os.path.isfile(xls_new_pdf_name):
        os.remove(xls_new_pdf_name)
    else:    ## Show an error ##
        print("Error: %s file not found" %  xls_new_pdf_name)


def gs_compress(pdf_final):
    # source_file = os.path.basename(pdf_final)
    # filename = pdf_final
    output_dir = os.path.dirname(pdf_final)
    # output_dir = r"c:\faturas"
    output_file_name = os.path.join(output_dir, "c_" + os.path.basename(pdf_final))
    output_file = '-sOutputFile=' + output_file_name

    args = ['C:\\Program Files\\gs\\gs10.00.0\\bin\\GSWIN64C.EXE',
            '-dPDFX',
            '-dBATCH',
            '-dNOPAUSE',
            '-dPDFSETTINGS=/ebook',
            '-dEmbedAllFonts=true',
            '-dSubsetFonts=true',
            '-dSimulateOverprint=true',
            '-sColorConversionStrategy=CMYK',
            '-sDEVICE=pdfwrite',
            '-dColorImageDownsampleType=/Bicubic',
            '-dColorImageResolution=150',
            '-dGrayImageDownsampleType=/Bicubic',
            '-dGrayImageResolution=150',
            '-dMonoImageDownsampleType=/Bicubic',
            '-dMonoImageResolution=150',
            '-dQUIET',
            str(output_file),
            pdf_final
            ]

    proc = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)

    # os.system("C:\\Program Files\\gs\\gs10.00.0\\bin\\GSWIN64C.EXE -dNOPAUSE -sDEVICE=jpeg -r144 -sOutputFile=" + output_file + ' ' + pdf_final)

    # -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/printer -dNOPAUSE -dQUIET -dBATCH -sOutputFile=output.pdf input.pdf
    # -q -dNOPAUSE -dBATCH -dSAFER -dSimulateOverprint=true -sDEVICE=pdfwrite -dPDFSETTINGS=/ebook -dEmbedAllFonts=true -dSubsetFonts=true -dAutoRotatePages=/None -dColorImageDownsampleType=/Bicubic -dColorImageResolution=150 -dGrayImageDownsampleType=/Bicubic -dGrayImageResolution=150 -dMonoImageDownsampleType=/Bicubic -dMonoImageResolution=150 -sOutputFile=output.pdf input.pdf

    # gswin64c.exe -dPDFSETTINGS#/ebook -dPDFX -dBATCH -dNOPAUSE -sColorConversionStrategy=CMYK -sDEVICE=pdfwrite -sOutputFile="output2.pdf" FRS-2065590673.pdf

    print(proc.communicate())

    print(f"Arquivo pdf compactado: {output_file_name}")



if __name__ == '__main__':

    parser = argparse.ArgumentParser()

    parser.add_argument('-f', '--fpath', help='files path', default="")
    parser.add_argument('-c', '--cpath', help='compress files path', default="")

    args = parser.parse_args()

    fpath = args.fpath

    if args.cpath:
        gs_compress(args.cpath)

    if not fpath:
        fpath = PDF_FOLDER

    if not os.path.isdir(fpath):
        print(f"Pasta origem inválida: {fpath}")
        sys.exit(-1)
    else:
        xlsx_files = glob.glob(fpath + "\\*.xlsx")
        pdf_files = glob.glob(fpath + "\\*.pdf")

    regex = re.compile(r'\d{5,15}')

    print(f"Using folder: {fpath}")

    pdf_to_join = []

    for xls_file in xlsx_files:

        xls_file_name = os.path.basename(xls_file)
        cod_fatura = regex.findall(xls_file_name)[0]
        pdf_files = glob.glob(fpath + f"\\*{cod_fatura}*.pdf")

        if not pdf_files:
            print(f"Sem pdf da conta correspondente para fatura {cod_fatura}")
            continue

        if len(pdf_files) > 1:
            print(f"Mais de um arquivo PDF com o codigo da fatura, deixe apenas so um na pasta: {cod_fatura}")
            sys.exit(-1)

        xls_new_pdf_name = "1." + xls_file_name.replace(".xlsx", ".xlsx.pdf")

        print(f"Convertendo {xls_file} para {xls_new_pdf_name}")
        xlsx_pdf(xls_file, os.path.join(fpath, xls_new_pdf_name))

        pdf_to_join.append(os.path.join(fpath, xls_new_pdf_name))
        pdf_to_join.append(pdf_files[0])

        pdf_final = os.path.join(fpath, f"{cod_fatura}.out.pdf")
        pdf_join(pdf_to_join, pdf_final)
        pdf_to_join = []

        gs_compress(pdf_final)

        #pdf_compress(pdf_final)

        delete_file(os.path.join(fpath, xls_new_pdf_name))

        delete_file(os.path.join(fpath, f"{cod_fatura}.out.pdf"))





