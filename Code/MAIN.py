import os
from Code.pdf_tools import excel_to_pdf, trace_dict_pagenumber, add_bookmarks, excel_to_pdf2
from Code.table_of_content import duplicator, generate_chapters, generate_paragraphs, properClosure, chap_and_par_combine, table_of_content_cleaner
from Code.create_A4_worksheet import create_A4_TOC_empty, update_A4_TOC
from Code.Move import moveSheet
from Code.deleter import deleteFiles
from Code.path_stuff import path_dissecting
from Code.sort import paragraph_sort
import time

report_selection = None
appendices_selection = None

# function number one;
# creates full pdf report from excel file including table of content to pdf and excel, and bookmarks.
def full_pdf_report(report, appendix, pathInp):
    global report_selection, appendices_selection
    report_selection = report
    appendices_selection = appendix

    start_time = time.time()

    '''
    More info on this tool is found in the READ_ME.txt file.
    Questions about the tool can be send to:
    dervis.vanleersum@gmail.com
    '''

    # dissecting the path and returning:
    #   - file name with extension
    #   - file directory
    #   - file name without extension
    stepZero = path_dissecting(pathInp)

    # from the directory as input this creates a duplicate file and returns:
    #   - the duplicate path
    #   - the duplicate directory
    stepOne = duplicator(stepZero[1], pathInp)

    # with duplicate path, and selected excel tabs as input this returns:
    #   - a dictionary with chapters as keys and empty values
    #   - pandas dataframe of the duplicate file
    stepTwo = generate_chapters(stepOne[0], report_selection)

    # with the pandas dataframe, and selected excel tabs as input this returns:
    #   - a lists for each sheet (chapter) in the duplicate file containing all the paragraphs
    stepThree = generate_paragraphs(stepTwo[1], report_selection)

    # with the chapters dictionary and lists of paragraphs as input this returns:
    #   - a dictionary with all chapters as keys and all paragraphs lists as values
    stepFour = chap_and_par_combine(stepTwo[0], stepThree)

    # with the pandas dataframe as input this function returns nothing but Closes the df file
    properClosure(stepTwo[1])

    # with the file directory and TOC dictionary as input, this function returns:
    #   - the file path of TOC.xlsx
    #   - the number of TOC pages
    # output:  this function generates an empty xlsx file for the table of content in specified layout
    stepFive = create_A4_TOC_empty(stepZero[1], stepFour)

    # with the Excel report file path and file directory as input this function returns:
    #   - the created pdf report file path
    # output:  function exports a pdf file from the defined Excel report workbook
    stepSix = excel_to_pdf(pathInp, stepZero[1], report_selection)

    # with chapPars dictionary, the PDF report file path and TOC amount of pages input this function returns:
    #   - a dictionary with all chapters as keys, all paragraphs lists as values and the pagenumber attached to all values
    stepSeven = trace_dict_pagenumber(stepFour, stepSix, stepFive[1])

    # with the chapPars dictionary from step seven this function returns:
    #   - a cleaned up dictionary where chapters and paragraphs without a pagenumber are deleted
    stepEight = table_of_content_cleaner(stepSeven)

    # with the chapPars dictionary from step eight this function returns:
    #   - a sorted dictionary where aragraphs are in the right order
    stepNine = paragraph_sort(stepEight)

    # with the cleaned and sorted chapPars dictionary from step eight this function returns nothing:
    # output:   function updates the TOC xlsx file created in step five
    update_A4_TOC(stepFive[0], stepNine, stepFive[1])

    # with the original file path and the TOC file path as input, this function does not return anything
    # the TOC is added to the Excel file
    moveSheet(pathInp, stepFive[0])

    # returns nothing. Function deletes the TOC.xlsx file and deletes the duplicate
    deleteFiles(stepSix)

    # with the Excel report file path and file directory as input this function returns:
    #   - the created pdf report file path
    # output:  function exports a pdf file from the defined Excel report workbook
    stepEleven = excel_to_pdf2(pathInp, stepZero[1], report_selection, stepFive[2])

    val = 0
    for file in os.listdir(stepZero[1]):
        if file.endswith(".pdf"):
            if stepZero[2] in file:
                val += 1
    
    final_path = stepZero[1] + "\\" + stepZero[2] + "_" + str(val) + ".pdf"

    # with the Excel report file path and file directory as input this function returns:
    #   - the created pdf report file path
    # output:  function exports a pdf file from the defined Excel report workbook
    add_bookmarks(stepEleven, stepNine, final_path)

    # returns nothing. Function deletes the TOC.xlsx file and deletes the duplicate
    deleteFiles(stepFive[0], stepOne[0], stepEleven)

    print("Program took", time.time() - start_time, "seconds to run.")


# function number two;
# creates table of content to excel file.
def toc_to_excel(report, appendix, pathInp):
    global report_selection, appendices_selection
    report_selection = report
    appendices_selection = appendix

    start_time = time.time()

    '''
    More info on this tool is found in the READ_ME.txt file.
    Questions about the tool can be send to:
    dervis.vanleersum@gmail.com
    '''
    # dissecting the path and returning:
    #   - file name with extension
    #   - file directory
    #   - file name without extension
    stepZero = path_dissecting(pathInp)

    # from the directory as input this creates a duplicate file and returns:
    #   - the duplicate path
    #   - the duplicate directory
    stepOne = duplicator(stepZero[1], pathInp)

    # with duplicate path, and selected excel tabs as input this returns:
    #   - a dictionary with chapters as keys and empty values
    #   - pandas dataframe of the duplicate file
    stepTwo = generate_chapters(stepOne[0], report_selection)

    # with the pandas dataframe, and selected excel tabs as input this returns:
    #   - a lists for each sheet (chapter) in the duplicate file containing all the paragraphs
    stepThree = generate_paragraphs(stepTwo[1], report_selection)

    # with the chapters dictionary and lists of paragraphs as input this returns:
    #   - a dictionary with all chapters as keys and all paragraphs lists as values
    stepFour = chap_and_par_combine(stepTwo[0], stepThree)

    # with the pandas dataframe as input this function returns nothing but Closes the df file
    properClosure(stepTwo[1])

    # with the file directory and TOC dictionary as input, this function returns:
    #   - the file path of TOC.xlsx
    #   - the number of TOC pages
    # output:  this function generates an empty xlsx file for the table of content in specified layout
    stepFive = create_A4_TOC_empty(stepZero[1], stepFour)

    # with the Excel report file path and file directory as input this function returns:
    #   - the created pdf report file path
    # output:  function exports a pdf file from the defined Excel report workbook
    stepSix = excel_to_pdf(pathInp, stepZero[1], report_selection)

    # with chapPars dictionary, the PDF report file path and TOC amount of pages input this function returns:
    #   - a dictionary with all chapters as keys, all paragraphs lists as values and the pagenumber attached to all values
    stepSeven = trace_dict_pagenumber(stepFour, stepSix, stepFive[1])

    # with the chapPars dictionary from step seven this function returns:
    #   - a cleaned up dictionary where chapters and paragraphs without a pagenumber are deleted
    stepEight = table_of_content_cleaner(stepSeven)

    # with the chapPars dictionary from step eight this function returns:
    #   - a sorted dictionary where aragraphs are in the right order
    stepNine = paragraph_sort(stepEight)

    # with the cleaned and sorted chapPars dictionary from step eight this function returns nothing:
    # output:   function updates the TOC xlsx file created in step five
    update_A4_TOC(stepFive[0], stepNine, stepFive[1])

    # with the original file path and the TOC file path as input, this function does not return anything
    # the TOC is added to the Excel file
    moveSheet(pathInp, stepFive[0])

    # returns nothing. Function deletes the TOC.xlsx file and deletes the duplicate
    deleteFiles(stepSix)

    # returns nothing. Function deletes the TOC.xlsx file and deletes the duplicate
    deleteFiles(stepFive[0], stepOne[0])

    print("Program took", time.time() - start_time, "seconds to run.")


# function number three;
# creates seperate pdf file from excel file including table of content.
def toc_sep_pdf(report, appendix, pathInp):
    global report_selection, appendices_selection
    report_selection = report
    appendices_selection = appendix

    start_time = time.time()

    '''
    More info on this tool is found in the READ_ME.txt file.
    Questions about the tool can be send to:
    dervis.vanleersum@gmail.com
    '''

    # dissecting the path and returning:
    #   - file name with extension
    #   - file directory
    #   - file name without extension
    stepZero = path_dissecting(pathInp)

    # from the directory as input this creates a duplicate file and returns:
    #   - the duplicate path
    #   - the duplicate directory
    stepOne = duplicator(stepZero[1], pathInp)

    # with duplicate path, and selected excel tabs as input this returns:
    #   - a dictionary with chapters as keys and empty values
    #   - pandas dataframe of the duplicate file
    stepTwo = generate_chapters(stepOne[0], report_selection)

    # with the pandas dataframe, and selected excel tabs as input this returns:
    #   - a lists for each sheet (chapter) in the duplicate file containing all the paragraphs
    stepThree = generate_paragraphs(stepTwo[1], report_selection)

    # with the chapters dictionary and lists of paragraphs as input this returns:
    #   - a dictionary with all chapters as keys and all paragraphs lists as values
    stepFour = chap_and_par_combine(stepTwo[0], stepThree)

    # with the pandas dataframe as input this function returns nothing but Closes the df file
    properClosure(stepTwo[1])

    # with the file directory and TOC dictionary as input, this function returns:
    #   - the file path of TOC.xlsx
    #   - the number of TOC pages
    # output:  this function generates an empty xlsx file for the table of content in specified layout
    stepFive = create_A4_TOC_empty(stepZero[1], stepFour)

    # with the Excel report file path and file directory as input this function returns:
    #   - the created pdf report file path
    # output:  function exports a pdf file from the defined Excel report workbook
    stepSix = excel_to_pdf(pathInp, stepZero[1], report_selection)

    # with chapPars dictionary, the PDF report file path and TOC amount of pages input this function returns:
    #   - a dictionary with all chapters as keys, all paragraphs lists as values and the pagenumber attached to all values
    stepSeven = trace_dict_pagenumber(stepFour, stepSix, stepFive[1])

    # with the chapPars dictionary from step seven this function returns:
    #   - a cleaned up dictionary where chapters and paragraphs without a pagenumber are deleted
    stepEight = table_of_content_cleaner(stepSeven)

    # with the chapPars dictionary from step eight this function returns:
    #   - a sorted dictionary where aragraphs are in the right order
    stepNine = paragraph_sort(stepEight)

    # with the cleaned and sorted chapPars dictionary from step eight this function returns nothing:
    # output:   function updates the TOC xlsx file created in step five
    update_A4_TOC(stepFive[0], stepNine, stepFive[1])

    # with the original file path and the TOC file path as input, this function does not return anything
    # the TOC is added to the original Excel file
    moveSheet(pathInp, stepFive[0])

    # returns nothing. Function deletes the duplicate
    deleteFiles(stepSix)

    # with the Excel report file path and file directory as input this function returns:
    #   - the created pdf report file path
    # output:  function exports a pdf file from the defined Excel report workbook
    stepEleven = excel_to_pdf2(stepFive[0], stepZero[1], ["TOC"], stepFive[2])

    # returns nothing. Function deletes the TOC.xlsx file and deletes the duplicate
    deleteFiles(stepFive[0], stepOne[0])

    print("Program took", time.time() - start_time, "seconds to run.")


# function number four;
# creates seperate pdf file from excel file including table of content.
def toc_sep_excel(report, appendix, pathInp):
    global report_selection, appendices_selection
    report_selection = report
    appendices_selection = appendix

    start_time = time.time()

    '''
    More info on this tool is found in the READ_ME.txt file.
    Questions about the tool can be send to:
    dervis.vanleersum@gmail.com
    '''

    # dissecting the path and returning:
    #   - file name with extension
    #   - file directory
    #   - file name without extension
    stepZero = path_dissecting(pathInp)

    # from the directory as input this creates a duplicate file and returns:
    #   - the duplicate path
    #   - the duplicate directory
    stepOne = duplicator(stepZero[1], pathInp)

    # with duplicate path, and selected excel tabs as input this returns:
    #   - a dictionary with chapters as keys and empty values
    #   - pandas dataframe of the duplicate file
    stepTwo = generate_chapters(stepOne[0], report_selection)

    # with the pandas dataframe, and selected excel tabs as input this returns:
    #   - a lists for each sheet (chapter) in the duplicate file containing all the paragraphs
    stepThree = generate_paragraphs(stepTwo[1], report_selection)

    # with the chapters dictionary and lists of paragraphs as input this returns:
    #   - a dictionary with all chapters as keys and all paragraphs lists as values
    stepFour = chap_and_par_combine(stepTwo[0], stepThree)

    # with the pandas dataframe as input this function returns nothing but Closes the df file
    properClosure(stepTwo[1])

    # with the file directory and TOC dictionary as input, this function returns:
    #   - the file path of TOC.xlsx
    #   - the number of TOC pages
    # output:  this function generates an empty xlsx file for the table of content in specified layout
    stepFive = create_A4_TOC_empty(stepZero[1], stepFour)

    # with the Excel report file path and file directory as input this function returns:
    #   - the created pdf report file path
    # output:  function exports a pdf file from the defined Excel report workbook
    stepSix = excel_to_pdf(pathInp, stepZero[1], report_selection)

    # with chapPars dictionary, the PDF report file path and TOC amount of pages input this function returns:
    #   - a dictionary with all chapters as keys, all paragraphs lists as values and the pagenumber attached to all values
    stepSeven = trace_dict_pagenumber(stepFour, stepSix, stepFive[1])

    # with the chapPars dictionary from step seven this function returns:
    #   - a cleaned up dictionary where chapters and paragraphs without a pagenumber are deleted
    stepEight = table_of_content_cleaner(stepSeven)

    # with the chapPars dictionary from step eight this function returns:
    #   - a sorted dictionary where aragraphs are in the right order
    stepNine = paragraph_sort(stepEight)

    # with the cleaned and sorted chapPars dictionary from step eight this function returns nothing:
    # output:   function updates the TOC xlsx file created in step five
    update_A4_TOC(stepFive[0], stepNine, stepFive[1])

    # returns nothing. Function deletes the TOC.xlsx file and deletes the duplicate
    deleteFiles(stepSix, stepOne[0])

    print("Program took", time.time() - start_time, "seconds to run.")


# function number five;
# creates pdf report from selected tabs in excel, with bookmarks.
def excel_to_pdfreport():
    return