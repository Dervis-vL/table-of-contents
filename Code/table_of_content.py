import pandas as pd
import os, re, shutil
from datetime import date


'''
This code will retrieve chapters and paragraphs from a report in Excel.
In order to work propperly, the following settings will need to be applied to the report.
'''

# get the full path of the file you want to create a table of content of
# pathInp = input('Insert full path with extension:\n')

# duplicate the file in question to the same directory to prevent corruption to original file
def duplicator(directory, path):
    dupli = 0
    for file in os.listdir(directory):
        if 'DUPLICATE_' in file.upper():
            dupli += 1
    
    today = date.today()
    d = today.strftime("%d-%m-%y")

    dupName = str(dupli) + "_duplicate_" + str(d) + ".xlsx"
    duplicate = directory + "\\" + dupName

    shutil.copy(path, duplicate)
    
    return duplicate, directory

# create a function that will gather all the chapters from an excel file
def generate_chapters(pathDuplicate, report_ws):
    chapPar = {}
    chaptersOptions = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"]

    # open and load the (duplicate) workbook
    df = pd.ExcelFile(pathDuplicate)

    tabs = []
    for itm in report_ws:
        tabs.append(str(itm))
    
    for sheet in df.sheet_names:
        if (sheet in tabs) and (sheet.upper()[0] != "C"):
            for i in range(0, 4):
                dfCol_A = pd.read_excel(df, header=None, sheet_name=sheet).iloc[i, 0]
                dfCol_B = pd.read_excel(df, header=None, sheet_name=sheet).iloc[i, 1]
                dfCol_C = pd.read_excel(df, header=None, sheet_name=sheet).iloc[i, 2]
                dfCol_D = pd.read_excel(df, header=None, sheet_name=sheet).iloc[i, 3]
                if (pd.isnull(dfCol_A) != True) and (pd.isnull(dfCol_B) != True) and (pd.isnull(dfCol_C) == True) and (str(dfCol_A)[0] in chaptersOptions):
                    chapPar.update({str(dfCol_A) + " " + str(dfCol_B):[]})
                    column = 'A'
                    break
                elif (pd.isnull(dfCol_A) != True) and (pd.isnull(dfCol_B) == True) and (pd.isnull(dfCol_C) != True) and (str(dfCol_A)[0] in chaptersOptions):
                    chapPar.update({str(dfCol_A) + " " + str(dfCol_C):[]})
                    column = 'A'
                    break
                elif (pd.isnull(dfCol_A) != True) and (pd.isnull(dfCol_B) == True) and (pd.isnull(dfCol_C) == True) and (str(dfCol_A)[0] in chaptersOptions):
                    chapPar.update({str(dfCol_A):[]})
                    column = 'A'
                    break
                elif (pd.isnull(dfCol_A) == True) and (pd.isnull(dfCol_B) != True) and (pd.isnull(dfCol_C) != True) and (str(dfCol_B)[0] in chaptersOptions):
                    chapPar.update({str(dfCol_B) + " " + str(dfCol_C):[]})
                    column = 'B'
                    break
                elif (pd.isnull(dfCol_A) == True) and (pd.isnull(dfCol_B) != True) and (pd.isnull(dfCol_C) == True) and (str(dfCol_B)[0] in chaptersOptions):
                    chapPar.update({str(dfCol_B):[]})
                    column = 'B'
                    break
                elif (pd.isnull(dfCol_A) == True) and (pd.isnull(dfCol_B) == True) and (pd.isnull(dfCol_C) != True) and (pd.isnull(dfCol_D) != True) and (str(dfCol_C)[0] in chaptersOptions):
                    chapPar.update({str(dfCol_C) + " " + str(dfCol_D):[]})
                    column = 'C'
                    break
                elif (pd.isnull(dfCol_A) == True) and (pd.isnull(dfCol_B) == True) and (pd.isnull(dfCol_C) != True) and (pd.isnull(dfCol_D) == True) and (str(dfCol_C)[0] in chaptersOptions):
                    chapPar.update({str(dfCol_C):[]})
                    column = 'C'
                    break
                elif (pd.isnull(dfCol_A) != True) and (pd.isnull(dfCol_B) != True) and (pd.isnull(dfCol_C) == True) and (str(dfCol_A)[:2] in chaptersOptions):
                    chapPar.update({str(dfCol_A) + " " + str(dfCol_B):[]})
                    column = 'A'
                    break
                elif (pd.isnull(dfCol_A) != True) and (pd.isnull(dfCol_B) == True) and (pd.isnull(dfCol_C) != True) and (str(dfCol_A)[:2] in chaptersOptions):
                    chapPar.update({str(dfCol_A) + " " + str(dfCol_C):[]})
                    column = 'A'
                    break
                elif (pd.isnull(dfCol_A) != True) and (pd.isnull(dfCol_B) == True) and (pd.isnull(dfCol_C) == True) and (str(dfCol_A)[:2] in chaptersOptions):
                    chapPar.update({str(dfCol_A):[]})
                    column = 'A'
                    break
                elif (pd.isnull(dfCol_A) == True) and (pd.isnull(dfCol_B) != True) and (pd.isnull(dfCol_C) != True) and (str(dfCol_B)[:2] in chaptersOptions):
                    chapPar.update({str(dfCol_B) + " " + str(dfCol_C):[]})
                    column = 'B'
                    break
                elif (pd.isnull(dfCol_A) == True) and (pd.isnull(dfCol_B) != True) and (pd.isnull(dfCol_C) == True) and (str(dfCol_B)[:2] in chaptersOptions):
                    chapPar.update({str(dfCol_B):[]})
                    column = 'B'
                    break
                elif (pd.isnull(dfCol_A) == True) and (pd.isnull(dfCol_B) == True) and (pd.isnull(dfCol_C) != True) and (pd.isnull(dfCol_D) != True) and (str(dfCol_C)[:2] in chaptersOptions):
                    chapPar.update({str(dfCol_C) + " " + str(dfCol_D):[]})
                    column = 'C'
                    break
                elif (pd.isnull(dfCol_A) == True) and (pd.isnull(dfCol_B) == True) and (pd.isnull(dfCol_C) != True) and (pd.isnull(dfCol_D) == True) and (str(dfCol_C)[:2] in chaptersOptions):
                    chapPar.update({str(dfCol_C):[]})
                    column = 'C'
                    break
                
    return chapPar, df


# function to obtain all the paragraphs from the duplicate workbook and stores the paragraphs of each chapter in seperate lists
def generate_paragraphs(df, report_ws):
    parList_1 = []
    parList_2 = []

    tabs = []
    for itm in report_ws:
        tabs.append(str(itm))

    for sheet in df.sheet_names:
        if sheet in tabs:
            dfRead = pd.read_excel(df, header=None, sheet_name=sheet)
            for x in range(0,(dfRead.shape[1]-6)):
                listed_1 = dfRead.iloc[:, x].tolist()
                listed_2 = dfRead.iloc[:, x+1].tolist()
                listed_3 = []
                for (itm_1, itm_2) in zip(listed_1, listed_2):
                    listed_3.append(str(itm_1) + "  " + str(itm_2))
                for itm in listed_1:
                    if re.search(r'^[1-9]\.[1-9]\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_1.append(str(itm))
                    elif re.search(r'^[1-9]\.[1-9]\d\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_1.append(str(itm))
                    elif re.search(r'^[1-2]\d\.[1-9]\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_1.append(str(itm))
                    elif re.search(r'^[1-2]\d\.[1-9]\d\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_1.append(str(itm))
                    elif re.search(r'^[1-9]\.[1-9]\.\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_1.append(str(itm))
                    elif re.search(r'^[1-9]\.[1-9]\d\.\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_1.append(str(itm))
                    elif re.search(r'^[1-2]\d\.[1-9]\.\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_1.append(str(itm))
                    elif re.search(r'^[1-2]\d\.[1-9]\d\.\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_1.append(str(itm))
                for itm in listed_3:
                    if re.search(r'^[1-9]\.[1-9]\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_2.append(str(itm))
                    elif re.search(r'^[1-9]\.[1-9]\d\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_2.append(str(itm))
                    elif re.search(r'^[1-2]\d\.[1-9]\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_2.append(str(itm))
                    elif re.search(r'^[1-2]\d\.[1-9]\d\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_2.append(str(itm))
                    elif re.search(r'^[1-9]\.[1-9]\.\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_2.append(str(itm))
                    elif re.search(r'^[1-9]\.[1-9]\d\.\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_2.append(str(itm))
                    elif re.search(r'^[1-2]\d\.[1-9]\.\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_2.append(str(itm))
                    elif re.search(r'^[1-2]\d\.[1-9]\d\.\s\s\D\D\D\D\D', str(itm)) and '=' not in itm:
                        parList_2.append(str(itm))

    totalPar = []
    if len(parList_1) == 0:
        for chap in tabs:
            parList = []
            for par in parList_2:
                if str(chap + '.') == str(par[:2]):
                    parList.append(par)
                elif str(chap) == str(par[:2]):
                    parList.append(par)
            totalPar.append(sorted(parList))
    else:
        for chap in tabs:
            parList = []
            for par in parList_1:
                if str(chap + '.') == str(par[:2]):
                    parList.append(par)
                elif str(chap) == str(par[:2]):
                    parList.append(par)
            totalPar.append(sorted(parList))
    return totalPar

# Function to get rid of all the evidence. Close workbook remove the duplicate and hide the body.
def properClosure(dfDuplicate):
    dfDuplicate.close()

# Function to match the dictionary values (lists of paragraphs) to the right keys (chapters)
def chap_and_par_combine(chap, par):
    for key in chap:
        for itm in par:
            if len(itm) > 0:
                if key[:2] == itm[0][:2]:
                    chap[key] = itm
                    par.remove(itm)
    for key in chap:
        for itm in par:
            if len(itm) > 0:
                if key[0] == itm[0][0]:
                    chap[key] = itm
                    par.remove(itm)
    return chap

# Function to clean up the table of content after adding pagenumbers to existing chapters and paragraphs.
            # last edited: 27-05-2021
def table_of_content_cleaner(chaps_and_pars):
    
    for chapter in [*chaps_and_pars.keys()]:
        try:
            page = chapter.split("_", 1)[1]
            if type(int(page)) == int:
                pass
        except Exception:
            chaps_and_pars.pop(chapter)
            continue

        fault_list = []
        for paragraph in chaps_and_pars.get(chapter):
            try:
                page = paragraph.split("_", 1)[1]
                if type(int(page)) == int:
                    continue
            except Exception:
                fault_list.append(paragraph)
                continue
        
        for fault in fault_list:
            chaps_and_pars.get(chapter).remove(fault)
    
    return chaps_and_pars


# if __name__ == __main__:
#     pass
    # stepOne = duplicator(pathInp)
    # stepTwo = generate_chapters(stepOne[0])
    # stepThree = generate_paragraphs(stepTwo[1])
    # stepFour = chap_and_par_combine(stepTwo[0], stepThree)
    # stepFive = properClosure(stepTwo[1])


    # print(stepFour)