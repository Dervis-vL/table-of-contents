# start with importing packages
import win32com.client, os, sys
import pandas as pd
from datetime import date
import PyPDF2, re
from PyPDF2 import PdfFileWriter, PdfFileReader

def excel_to_pdf(filePath, fileDirectory, tabs):
    '''creates pdf files from all sheets until sheet "COL"
    input:  full path of xlsx file, the directory of the xlsx file
    return: -
    output: pdf file of all worksheet until the colophon sheet


    safety: - checks if the report has a colophon (otherwise sys.exit())
            - numbers files to prevent overwriting existing pdf's
            - date stamps the pdf file
    '''
    # df = pd.ExcelFile(filePath)
    # ws_list = df.sheet_names

    # ws_print_list = []
    # for sheet in ws_list:
    #     if sheet in tabs:
    #         ws_print_list.append(sheet)

    # df.close()

    duplicate = 0
    for file in os.listdir(fileDirectory):
        if "COMBINED_" in file.upper():
            duplicate += 1

    today = date.today()
    d = today.strftime("%d-%m-%y")

    fileCombined = str(duplicate) + "_Combined_" + str(d) + ".pdf"
    path_to_pdf = fileDirectory + "/" + fileCombined

    path_to_pdf = path_to_pdf.replace("/", "\\")

    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = False
    o.DisplayAlerts = False
    wb = o.Workbooks.OpenXML(filePath)

    wb.WorkSheets(tabs).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf, Quality=0, IncludeDocProperties=True, IgnorePrintAreas=False, OpenAfterPublish=False)

    o.Quit()

    return path_to_pdf


def trace_dict_pagenumber(chapPars, path_to_pdf, pages):
    # open the pdf file
    object = open(path_to_pdf, 'rb')

    pdfReader = PyPDF2.PdfFileReader(object)
    
    # get number of pages
    numPages = pdfReader.getNumPages()
    
    # define keyterms
    chapters = [*chapPars.keys()]

    # frontpage is always assumed to be 2 pages TODO: add to read_me
    frontPages = 1  

    par_len = 0  

    # extract text and do the search
    for i in range(frontPages, int(numPages)):
        pageObj = pdfReader.getPage(i)
        text = pageObj.extractText()

        length_chap = len(chapters)
        if length_chap > 0:
            for chap_num in range(0, length_chap):
                chapter = chapters[chap_num]
            
                chap_start = [*chapPars.keys()][chap_num].split(" ", 1)[0]
                chap_end = [*chapPars.keys()][chap_num].split(" ", 1)[1].lstrip()

                regex_1 = re.search(f"{chap_start}\s*{chap_end}", text, re.MULTILINE)
                regex_2 = re.search(f"{chap_end}\s*{chap_start}", text, re.MULTILINE)
                if regex_1 or regex_2:
                    chapter_update = chapter + "_" + str(i+1+int(pages))
                    chapPars[chapter_update] = chapPars.pop(chapter)
                    chapters.pop(chap_num)

                    paragraphs = chapPars.get(chapter_update)
                    par_len = len(paragraphs)
                    break

        if par_len > 0:
            for par_num in range(0, par_len):
                par_start = chapPars.get(chapter_update)[par_num].split(" ", 1)[0]
                par_end = chapPars.get(chapter_update)[par_num].split(" ", 1)[1].lstrip()

                regex_3 = re.search(f"{par_start}\s*{par_end}", text, re.MULTILINE)
                regex_4 = re.search(f"{par_end}\s*{par_start}", text, re.MULTILINE)
                if regex_3 or regex_4:
                    paragraph = chapPars[chapter_update][par_num]
                    chapPars[chapter_update][par_num] = paragraph + "_" + str(i+1+int(pages))
        
    object.close()
    return chapPars


# winner for looping over each page and bookmark
def add_bookmarks(source, dictionary, pdf):
    chap_pages = [int(x.split("_", 1)[1]) for x in dictionary.keys()]  # extract page number from dictionary
    chapters = [x.split("_", 1)[0] for x in dictionary.keys()]         # extract chapter from dictionary
    # paragraphs = [x for x in dictionary.values()]

    pdf_in_file = open(source, 'rb')

    inputpdf = PdfFileReader(pdf_in_file)
    pages_no = inputpdf.numPages

    output = PdfFileWriter()

    # Set bookmark for first page
    output.addPage(inputpdf.getPage(0))                     # extract first page to set bookmark for first page
    output.addBookmark('Voorpagina', 0, parent=None)    # set bookmark on extracted page
    # output.setPageMode("/UseOutlines")
    
    for i in range(1, pages_no):
        output.addPage(inputpdf.getPage(i))
        count = 0
        for x in chap_pages:
            if (i+1) == x:
                parent = output.addBookmark(chapters[count], i, parent=None)
                # output.setPageMode("/UseOutlines")
                
            # parent_paragraphs = paragraphs[count]
            # parent_par_pages = [int(x.split("_", 1)[1]) for x in parent_paragraphs]
            # parent_par = [x.split("_", 1)[0] for x in parent_paragraphs]

            # loop_par = 0
            # for y in parent_par_pages:
            #     if (i+1) == y:
            #         output.addBookmark(parent_par[loop_par], i, parent)
            #         # output.setPageMode("/UseOutlines")
            #     loop_par += 1

            count += 1

    outputStream = open(pdf, 'wb')

    output.write(outputStream)

    pdf_in_file.close()
    outputStream.close()


def excel_to_pdf2(filePath, fileDirectory, tabs, title_TOC):
    '''creates pdf files from all sheets until sheet "COL"
    input:  full path of xlsx file, the directory of the xlsx file
    return: -
    output: pdf file of all worksheet until the colophon sheet


    safety: - checks if the report has a colophon (otherwise sys.exit())
            - numbers files to prevent overwriting existing pdf's
            - date stamps the pdf file
    '''
    # df = pd.ExcelFile(filePath)
    # ws_list = df.sheet_names

    # ws_print_list = []
    # for sheet in ws_list:
    #     if sheet in tabs:
    #         ws_print_list.append(sheet)

    # df.close()

    duplicate = 0
    if len(tabs) > 1:
        for file in os.listdir(fileDirectory):
            if "COMBINED_" in file.upper():
                duplicate += 1
    else:
        for file in os.listdir(fileDirectory):
            if "TOC_" in file.upper():
                duplicate += 1

    today = date.today()
    d = today.strftime("%d-%m-%y")

    if len(tabs) > 1:
        fileCombined = str(duplicate) + "_Combined_" + str(d) + ".pdf"
        path_to_pdf = fileDirectory + "/" + fileCombined
    else:
        fileCombined = str(duplicate) + "_TOC_" + str(d) + ".pdf"
        path_to_pdf = fileDirectory + "/" + fileCombined

    path_to_pdf = path_to_pdf.replace("/", "\\")

    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = False
    o.DisplayAlerts = False
    wb = o.Workbooks.OpenXML(filePath)
    
    if len(tabs) > 1:
        # tabs[1] = [title_TOC]
        tabs = tabs[:1] + [title_TOC] + tabs[1:]

    wb.WorkSheets(tabs).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf, Quality=0, IncludeDocProperties=True, IgnorePrintAreas=False, OpenAfterPublish=False)

    o.Quit()

    return path_to_pdf
