import os

def pathDetector():
    pathInp = input('Insert full path with extension:\n')

    if pathInp.startswith("'") and pathInp.endswith("'"):
        pathInp = pathInp[1:-1]
        return pathInp
    elif pathInp.startswith('"') and pathInp.endswith('"'):
        pathInp = pathInp[1:-1]
        return pathInp
    else:
        return pathInp


def path_dissecting(inp_path):
    # obtain filename with extension
    fileName_type = os.path.basename(inp_path)
    
    # obtain directory from path variable
    directory = os.path.dirname(inp_path)

    # obtain filename without extension
    fileName = os.path.splitext(fileName_type)[0]

    return fileName_type, directory, fileName

if __name__ == "__main__":
    path = r"C:\Users\ddvanleersum\Documents\01_Dervis\06_Programming\Python\Scripts\TOC_generator\Example reports\Excel\INFR180339-62B-123-01 - Op Gen Heck_V2.xlsx"

    test = path_dissecting(path)

    print(test[0])
    print(test[1])
    print(test[2])