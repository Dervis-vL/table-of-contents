from win32com.client import Dispatch

def moveSheet(dest, source):
    '''
    The purpose of this file is to copy excel worksheets with 
    index x from file A to index y of file B.

    All content and characteristics are copied.
    For now this function does not copy the print area to the new file.
    '''
    xl = Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    x = 1
    y = 2
    
    wbDest = xl.Workbooks.OpenXML(Filename=dest)

    title ='TOC'
    index = 1
    sheet_names = [sheet.Name for sheet in wbDest.Sheets]

    

    for sheet in sheet_names:
        if sheet == title:
            wbDest.Worksheets(index).Delete()
        index += 1

    wbSource = xl.Workbooks.OpenXML(Filename=source)

    wsSource = wbSource.Worksheets(x)
    wsSource.Copy(Before=wbDest.Worksheets(y))

    wbDest.Close(SaveChanges=True)
    wbSource.Close(SaveChanges=False)
    xl.Quit()