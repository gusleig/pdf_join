# XLSX PDF Joiner

This little humble script joins an Excel file with a PDF.

The CLI uses one parameter:

```
pdf-xlsx.py -f C:\my_files
```

1. Look for xlsx file 
2. Convert it to PDF
3. Join this new PDF with other in folder.
4. Compress using ghostscript

## Create the exe

```b
pyinstaller xls2pdf.py --onefile --noconsole
```
The exe will be in the dist folder.

ps: If windows or anti-virus flags the exe as malware, remove the --onefile