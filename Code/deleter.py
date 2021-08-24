import os
import pandas as pd

def deleteFiles(*args):
    for arg in args:
        os.remove(arg)

# this function will check if the TOC already exists before moving a new TOC worksheet
#   input   path of original file
def check_for_TOC(source):

    df = pd.ExcelFile(source)
    
