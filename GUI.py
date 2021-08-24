from tkinter import *
from PIL import ImageTk, Image
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import os
from Code import MAIN

# set root window images and icon
icoPath = r"C:\Users\ddvanleersum\Documents\01_Dervis\06_Programming\Python\Apps\TOC_v2\Icons\dervis_face.ico"
pngPath = r"C:\Users\ddvanleersum\Documents\01_Dervis\06_Programming\Python\Apps\TOC_v2\Icons\dervis_arms.png"

root = Tk()

# define the name of the root window that will appear
root.title("Excel report tool        ᕦ໒( ՞ ◡ ՞ )७ᕤ")
root.iconbitmap(icoPath)
# root.geometry("400x400+80+80")
headerLogo = ImageTk.PhotoImage(Image.open(pngPath))
header_label = Label(root, image=headerLogo)
header_label.grid(row=0, column=1)

# Shows developer 
status = Label(root, text="Made by Dervis van Leersum", relief=SUNKEN, anchor=E)
status.grid(row=6, column=0, columnspan=3, sticky=W+E)
status.config(font=("Arial", 8, "italic"))

# creates first top frame
frame_one = LabelFrame(root, text="Select action", padx=20, pady=5)
frame_one.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky=W+E)
frame_one.config(font=("Arial", 11))
frame_two = LabelFrame(root, text="Selected file", padx=20, pady=5)
frame_two.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky=W+E)
frame_two.config(font=("Arial", 11))

# show possible main window actions
actions = ["Full PDF report", "Table of Content to Excel report","Seperate PDF Table of Content" ,"Seperate Excel Table of Content"]

# creates radio buttons
action = IntVar()

radio_1 = Radiobutton(frame_one, text=actions[0], variable=action, value=1)
radio_1.grid(row=0, column=0, sticky=W)
radio_2 = Radiobutton(frame_one, text=actions[1], variable=action, value=2)
radio_2.grid(row=1, column=0, sticky=W)
radio_3 = Radiobutton(frame_one, text=actions[2], variable=action, value=3)
radio_3.grid(row=2, column=0, sticky=W)
radio_4 = Radiobutton(frame_one, text=actions[3], variable=action, value=4)
radio_4.grid(row=3, column=0, sticky=W)

# pop up dialog box
def dialog_box():
    global root_filename
    root.filename = filedialog.askopenfilename(initialdir=r"c:\Users", title="S elect Excel report", filetypes=(("Excel files", "*.xlsx"), ("Macro files", "*.xlsm")))
    name = os.path.basename(str(root.filename))
    name = name.replace("'mode='r' encoding='UTF-8'>", "")
    lbl_filename = Label(frame_two, text=name)
    lbl_filename.grid(row=0, column=1)
    root_filename = root.filename

    btn_start = Button(root, text="NEXT \u279E", command=lambda: start(root.filename), bd=2)
    btn_start.grid(row=4, column=1, pady=(0, 15))
    btn_start.config(font=("Arial", 10, "bold"))

# after three windows and input is given, this function runs the action in the MAIN.py file
def run(action):
    top_2.destroy()

    appendix_ws = []
    for intvar in selected_2:
        appendix_ws.append(intvar.get())

    report_ws_lst = []
    appendix_ws_lst = []

    for pos in range(len(report_ws)):
        if report_ws[pos] == 1:
            report_ws_lst.append(worksheets[pos])

    for pos in range(len(appendix_ws)):
        if appendix_ws[pos] == 1:
            appendix_ws_lst.append(worksheets[pos])

    if action.get() == 1:
        last_lvl = Toplevel()
        last_lvl.title("Busy.....")
        last_lvl.iconbitmap(icoPath)

        header_label = Label(last_lvl, image=headerLogo)
        header_label.grid(row=0, column=1)

        waiting = Label(last_lvl, text="I'm busy, just wait......")
        waiting.grid(row=1, column=1, padx=40, pady=15, sticky=W)
        waiting.config(font=("Arial", 11))

        MAIN.full_pdf_report(report_ws_lst, appendix_ws_lst, root_filename)

        last_lvl.destroy
        root.destroy
    
    elif action.get() == 2:
        last_lvl = Toplevel()
        last_lvl.title("Busy.....")
        last_lvl.iconbitmap(icoPath)

        header_label = Label(last_lvl, image=headerLogo)
        header_label.grid(row=0, column=1)

        waiting = Label(last_lvl, text="I'm busy, just wait......")
        waiting.grid(row=1, column=1, padx=40, pady=15, sticky=W)
        waiting.config(font=("Arial", 11))

        MAIN.toc_to_excel(report_ws_lst, appendix_ws_lst, root_filename)

        root.destroy

    elif action.get() == 3:
        last_lvl = Toplevel()
        last_lvl.title("Busy.....")
        last_lvl.iconbitmap(icoPath)

        header_label = Label(last_lvl, image=headerLogo)
        header_label.grid(row=0, column=1)

        waiting = Label(last_lvl, text="I'm busy, just wait......")
        waiting.grid(row=1, column=1, padx=40, pady=15, sticky=W)
        waiting.config(font=("Arial", 11))

        MAIN.toc_sep_pdf(report_ws_lst, appendix_ws_lst, root_filename)

        root.destroy

    elif action.get() == 4:
        last_lvl = Toplevel()
        last_lvl.title("Busy.....")
        last_lvl.iconbitmap(icoPath)

        header_label = Label(last_lvl, image=headerLogo)
        header_label.grid(row=0, column=1)

        waiting = Label(last_lvl, text="I'm busy, just wait......")
        waiting.grid(row=1, column=1, padx=40, pady=15, sticky=W)
        waiting.config(font=("Arial", 11))

        MAIN.toc_sep_excel(report_ws_lst, appendix_ws_lst, root_filename)

        root.destroy

# continues to third and final window
def middle(filename):
    global top_2, report_ws, selected_2
    top_1.destroy()

    selected_2 = []
    report_ws = []
    for intvar in selected_1:
        report_ws.append(intvar.get())
    top_2 = Toplevel()
    top_2.title("Select appedices")
    top_2.iconbitmap(icoPath)
    # top_2.geometry("300x800")

    frame_top2 = LabelFrame(top_2)
    frame_top2.grid(row=0, column=0, sticky=N, padx=15, pady=(15, 10))

    lbl_top2 = Label(frame_top2, text="Check ONLY the boxes of tabs that are added to the appendix.\nThis means excluding all other shit.", justify=LEFT)
    lbl_top2.grid(row=0, column=0, padx=8, pady=(5, 5), sticky=W)
    lbl_top2.config(font=("Arial", 11))

    for x in range(len(worksheets)):
        var = IntVar()
        Checkbutton(top_2, text=worksheets[x], variable=var).grid(row=x+1, column=0, sticky=W, padx=(80, 50))
        selected_2.append(var)

    btn_next_2 = Button(top_2, text="NEXT \u279E", command=lambda: run(action), bd=1)
    btn_next_2.grid(row=(len(worksheets)+2), column=0, pady=15)
    btn_next_2.config(font=("Arial", 10, "bold"))


# check if first window input is correct and continue to next window
def start(filename):
    global selected_1, top_1, worksheets
    selected_1 = []
    if (action.get() == 1) or (action.get() == 2) or (action.get() == 3) or (action.get() == 4):
        top_1 = Toplevel()
        top_1.title("Select main report")
        top_1.iconbitmap(icoPath)
        # top_1.geometry("300x800")

        frame_top1 = LabelFrame(top_1)
        frame_top1.grid(row=0, column=0, sticky=N, padx=15, pady=(15, 10))

        lbl_top1 = Label(frame_top1, text="Check ONLY the boxes of tabs, that create the main report.\nThis means excluding all appendices and shit.", justify=LEFT)
        lbl_top1.grid(row=0, column=0, padx=8, pady=(5, 5), sticky=W)
        lbl_top1.config(font=("Arial", 11))

        df = pd.ExcelFile(filename)
        worksheets = df.sheet_names
        df.close()

        for x in range(len(worksheets)):
            var = IntVar()
            Checkbutton(top_1, text=worksheets[x], variable=var).grid(row=x+1, column=0, sticky=W, padx=(80, 50))
            selected_1.append(var)
    else:
        response = messagebox.showinfo("Empty radio buttons", "No action selection found.\nSelect an option before proceeding.")
        Label(root, text=response)

    btn_next_1 = Button(top_1, text="NEXT \u279E", command=lambda: middle(filename), bd=1)
    btn_next_1.grid(row=(len(worksheets)+2), column=0, pady=15)
    btn_next_1.config(font=("Arial", 10, "bold"))

# creates buttons
btn_fileSelect = Button(root, text="Select file", command=dialog_box, borderwidth=3)
btn_fileSelect.grid(row=2, column=1, pady=(0, 10))
btn_fileSelect.config(font=("Arial", 10, "bold"))

btn_start = Button(root, text="NEXT \u279E", command=lambda: start(root_filename), bd=2, state=DISABLED)
btn_start.grid(row=4, column=1, pady=(0, 15))
btn_start.config(font=("Arial", 10, "bold"))

mainloop()


#   ['VB', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', 'col']
#   ['VB', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', 'col']