import tempfile
import sys
import os
import shutil
import getopt
import zipfile
import glob
import re
import tkinter
import tkinter.filedialog

def main(argv):
    inputFile = ''
    outputFile = ''
    tmpDir = ''
    zipPath = ''
    zipModPath = ''
    zipDir = ''

    if len(argv) <= 0:
        tkinter.Tk().withdraw()
        inputFile = tkinter.filedialog.askopenfilename(title="Select input file", filetypes=(('xlsx files', '*.xlsx'), ('xls files', '*.xls')))
        inputFileName = os.path.splitext(inputFile)[0]
        outputFile = tkinter.filedialog.asksaveasfilename(title="Save unprotected excel file", defaultextension='xlsx', filetypes=(('xlsx files', '*.xlsx'), ('xls files', '*.xls')), initialfile=inputFileName+'_nopwd')

    try:
        opts, args = getopt.getopt(argv, "hi:o:")
    except getopt.GetoptError:
        print("excelUnprotect.py -i <inputfile> -o <outputfile>")
        sys.exit(2)

    for opt, arg in opts:
        if opt == '-h':
            print("excelUnprotect.py -i <inputfile> -o <outputfile>")
            sys.exit()
        elif opt == '-i':
            inputFile = arg
        elif opt == '-o':
            outputFile = arg
    
    if not os.path.isfile(inputFile):
        print("Error: input file does not exists or specified path is not a file")
        sys.exit(1)
    if not os.path.splitext(inputFile)[1] in ('.xls', '.xlsx'):
        print("Error: input file is not an MS Excel Workbook")
        sys.exit(1)
    if not os.path.isdir(os.path.dirname(os.path.abspath(outputFile))):
        print("Error: output file path does not exists")
        sys.exit(1)

    with tempfile.TemporaryDirectory() as tmpDir:
        zipPath = os.path.join(tmpDir, 'workbook.zip')
        zipModPath = os.path.join(tmpDir, 'workbook_mod')
        zipDir = os.path.join(tmpDir, 'workbook')
        shutil.copyfile(inputFile, zipPath)
        with zipfile.ZipFile(zipPath, 'r') as zipObj:
            zipObj.extractall(zipDir)

        workbook = os.path.join(zipDir, 'xl', 'workbook.xml')
        worksheets = glob.glob(os.path.join(zipDir, 'xl', 'worksheets', 'sheet*.xml'))
        
        with open(workbook, 'r') as wb:
            wbData = wb.read()
        wbData = re.sub(r'<workbookProtection.*?/>', '', wbData)
        with open(workbook, 'w') as wb:
            wb.write(wbData)

        for worksheet in worksheets:
            with open(worksheet, 'r') as ws:
                wsData = ws.read()
            wsData = re.sub(r'<sheetProtection.*?/>', '', wsData)
            with open(worksheet, 'w') as ws:
                ws.write(wsData)
        
        shutil.make_archive(zipModPath, 'zip', zipDir)
        shutil.copyfile(zipModPath+'.zip', outputFile)

if __name__ == "__main__":
    main(sys.argv[1:])