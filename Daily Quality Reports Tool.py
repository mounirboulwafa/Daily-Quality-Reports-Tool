import PyPDF2
import tkinter
import os.path
import xlwt
import xlrd
import webbrowser
from xlwt import Workbook
from xlutils.copy import copy
from win32com.client import Dispatch
from tkinter import filedialog
from tkinter import messagebox
from tkinter import *
from os import path
import datetime

#######################
# Open file Dialog to choose the PDF Document

window = tkinter.Tk()
window.title(" ATLAS BERRY FARMS | Daily Quality Reports Tool")
window.geometry("500x400")
window.iconbitmap(default='img\\icon2.ico')
window.resizable(0, 0)


#######################
# Open contact window

def show_contact():
    win = Toplevel(window)
    win.title(" Contact")
    win.geometry("400x210")
    win.resizable(0, 0)

    title = tkinter.Label(win, text="Mounir Boulwafa\n\nÉlève Ingénieur\n"
                                    "Ingénierie Informatique, Big Data et Cloud Computing\nENSET MOHAMMEDIA")
    title.pack(ipadx=80, ipady=15, fill='both')

    l = tkinter.Label(win, text="Tel : (+212) 6 33 62 38 73\nEmail : mounirboulwafa@gmail.com\n")
    l.pack()

    b = tkinter.Button(win, text="Fermer", command=win.destroy)
    b.pack(pady=7, padx=80, ipadx=10)

    l1 = tkinter.Label(win, text="", bg="chartreuse4", width=300, height=1, )
    l1.pack()

    win.wm_attributes("-topmost", 1)
    win.grab_set()


#######################
# Open about window

def show_about():
    win = Toplevel(window)
    win.title(" About")
    win.geometry("400x200")
    win.resizable(0, 0)

    title = tkinter.Label(win, text="ATLAS BERRY FARMS\nDaily Quality Reports Tool v1.5",
                          fg="chartreuse4", font="Helvetica 12 bold")
    title.pack(ipadx=80, ipady=20, fill='both')

    l = tkinter.Label(win, text="Copyright © 2019. All rights reserved. \nby Mounir Boulwafa.\n")
    l.pack(ipadx=8, ipady=6, fill='both')

    b = tkinter.Button(win, text="Fermer", command=win.destroy)
    b.pack(pady=10, padx=10, ipadx=20)

    l1 = tkinter.Label(win, text="", bg="chartreuse4", width=300, height=1, )
    l1.pack()

    # set always on top
    win.attributes("-topmost", 1)
    win.grab_set()


def show_processing_end():
    messagebox.showinfo(" Processing", " Process completed successfully")


def show_processing_error():
    messagebox.showwarning(" Error : Process could not be completed !", " Maybe The Excel file is already open ! \n\n"
                                                                        "Or contact the developer to fix the bug !!")


logo = tkinter.PhotoImage(file="img\\logo.png")

logo_label = tkinter.Label(window, image=logo, width=500, height=110, )
logo_label.place(x=20, y=60)
logo_label.pack()

title_label = tkinter.Label(window, text="Daily Quality Reports Tool",
                            fg="white", bg="chartreuse4", width=300, height=2, font="Helvetica 16 bold").pack()

space = tkinter.Label(window, height=2, ).pack()

myPDF = None
excelFile = ""

#######################
# styles

style1 = xlwt.easyxf(
    'pattern: pattern solid, fore_colour green;'
    'font: colour white,height 260, bold True;'
    'align: horiz center')

style = xlwt.easyxf(
    'pattern: pattern solid, fore_colour green;'
    'font: colour white,height 220, bold True;'
    'align: horiz center')


def delete_worksheet(w_sheet):
    # w_sheet.write(3, 3, "")
    # w_sheet.write(4, 3, "")
    for x in range(0, 11):
        for y in range(1, 100):
            w_sheet.write(y, x, "")


def write_headers(w_sheet):
    # w_sheet.write(0, 3, 'Daily Quality Report', style)
    # w_sheet.write_merge(0, 0, 3, 4, 'Daily Quality Report', style1)

    w_sheet.write(0, 0, 'Grower receipt', style)
    w_sheet.write(0, 1, 'Item number', style)
    w_sheet.write(0, 2, 'Id Bloc', style)
    w_sheet.write(0, 3, 'Batch number', style)
    w_sheet.write(0, 4, 'Arrival date / QC check', style)
    w_sheet.write(0, 5, 'Quantity', style)
    w_sheet.write(0, 6, 'Variety', style)
    w_sheet.write(0, 7, 'Quantity in KGs', style)
    w_sheet.write(0, 8, 'Final Grading', style)
    w_sheet.write(0, 9, 'Final PFQ score', style)
    w_sheet.write(0, 10, 'Grower', style)
    w_sheet.write(0, 11, 'Ranch', style)

    #######################
    # Styling ( Column width )

    w_sheet.col(0).width = 4200
    w_sheet.col(1).width = 4000
    w_sheet.col(2).width = 3000
    w_sheet.col(3).width = 6000
    w_sheet.col(4).width = 6400
    w_sheet.col(5).width = 3400
    w_sheet.col(6).width = 3300
    w_sheet.col(7).width = 4500
    w_sheet.col(8).width = 4000
    w_sheet.col(9).width = 4800
    w_sheet.col(10).width = 3000
    w_sheet.col(11).width = 3000


def load_pdf():
    myPDF = filedialog.askopenfilename(initialdir="C:\\Users\\" + str(os.getlogin()) + " \\Desktop",
                                       title=" Select file",
                                       filetypes=(("PDF Documents", "*.pdf"), ("all files", "*.*")))

    # print("File is loaded")

    #######################
    # Read the PDF Document

    try:
        pdfFileObj = open(myPDF, 'rb')
        # pdfFileObj = open('pdf.pdf', 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        pages = pdfReader.numPages

        #######################
        # Save the Data As Excel File

        excelFile = str(myPDF).replace('.pdf', '.xls')
        # print("1111" + str(excelFile))

        #######################
        # Create the Excel File

        # if the Excel file exist : update
        if path.exists(excelFile):
            rb = xlrd.open_workbook(excelFile, formatting_info=1)
            wb = copy(rb)
            w_sheet = wb.get_sheet(0)
            #
            # vaaaa = w_sheet.sheet_names()
            # print("000000" + str(vaaaa))

            delete_worksheet(w_sheet)
            # print("File Updated")

        # if the Excel file doesn't exist : create new one
        else:
            wb = Workbook()
            w_sheet = wb.add_sheet('Daily Quality Report')

            write_headers(w_sheet)

            wb.save(excelFile)

        ###################
        # start processing

        # Inserting the logo
        # w_sheet.insert_bitmap('img\\logo.bmp', 0, 0, 2, 2)

        n1, n2, n3, n4, n5, n6, n7, n8, n9, n10 = 0

        # Reading GrowerReceipt & Ranche

        pageObj = pdfReader.getPage(0)
        pageOne = pageObj.extractText()
        # print(pageOne)

        Grower_Regex = r"Grower:.*\n(\w+)|Producteur:.*\n(\w+)"
        Ranch_Regex = r"Ranch:.*\n(\w+)|Ferme:.*\n(\w+)"

        Growers = re.finditer(Grower_Regex, pageOne)
        Ranches = re.finditer(Ranch_Regex, pageOne)

        for matchNum, match in enumerate(Growers, start=1):
            for groupNum in range(0, len(match.groups())):
                groupNum = groupNum + 1
                if match.group(groupNum) is not None:
                    Grower_val = int(match.group(groupNum))

        for matchNum, match in enumerate(Ranches, start=1):
            for groupNum in range(0, len(match.groups())):
                groupNum = groupNum + 1
                if match.group(groupNum) is not None:
                    Ranche_val = int(match.group(groupNum))

        # print(Grower_val)
        # print(Ranche_val)

        for x in range(0, pages):
            # print("\n--------- Page " + str(x + 1) + " ----------")
            pageObj = pdfReader.getPage(x)
            text = pageObj.extractText()
            # print(text)

            ######################
            #  RegExpressions

            GrowerReceipt_Regex = r"Grower receipt:.*\n(.*)|Bon de réception:.*\n(.*)"
            ItemNumber_Regex = r"Production method.*\n.*.*\n(.*)|Méthode de Production.*\n.*.*\n(.*)"
            IdBloc_Regex = r"Batch number:.*\n.*(.{8}).{2}|Numéro de Lot:.*\n.*(.{8}).{2}"
            BatchNumber_Regex = r"Batch number:.*\n(.*)|Numéro de Lot:.*\n(.*)"
            QC_Regex = r"QC check:.*\n(.*)|Contrôle qualité:.*\n(.*)"
            Quantity_Regex = r"(.*\n.*)MA MOU|(.*\n.*)MA DAC|(.*\n.*)MA LAR"
            Variety_Regex = r"MA MOU.*\n(.*)|MA DAC.*\n(.*)|MA LAR.*\n(.*)"
            KG_Regex = r"Quantity in KGs:.*\n(.*)|Quantité en KG:.*\n(.*)"
            Grading_Regex = r"Final Grading:.*\n(.*)|Classification  finale:.*\n(.*)"
            PFQ_score_Regex = r"Final PFQ score:.*\n(.*)|Score PFQ final:.*\n(.*)"

            ######################
            #

            GrowerReceipts = re.finditer(GrowerReceipt_Regex, text)
            QCs = re.finditer(QC_Regex, text)
            IdBlocs = re.finditer(IdBloc_Regex, text)
            BatchNumbers = re.finditer(BatchNumber_Regex, text)

            ItemNumbers = re.finditer(ItemNumber_Regex, text)
            Quantities = re.finditer(Quantity_Regex, text)
            Varieties = re.finditer(Variety_Regex, text)

            KGs = re.finditer(KG_Regex, text)
            Gradings = re.finditer(Grading_Regex, text)
            PFQ_scores = re.finditer(PFQ_score_Regex, text)

            ######################
            #

            n = 0

            n1 = n + n1
            for matchNum, match in enumerate(GrowerReceipts, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n1 + 1, 0, val)
                            w_sheet.write(n1 + 1, 10, Grower_val)
                            w_sheet.write(n1 + 1, 11, Ranche_val)

                n1 += 1

            n2 = n + n2
            for matchNum, match in enumerate(QCs, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = match.group(groupNum).rstrip()
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n2 + 1, 4, val)
                n2 += 1

            n10 = n + n10
            for matchNum, match in enumerate(BatchNumbers, start=1):
                # print("111111111")
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n10 + 1, 3, val)
                n10 += 1

            n3 = n + n3
            for matchNum, match in enumerate(IdBlocs, start=1):
                # print("111111111")
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n3 + 1, 2, val)
                n3 += 1

            n4 = n + n4
            for matchNum, match in enumerate(ItemNumbers, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n4 + 1, 1, int(val))
                n4 += 1

            n5 = n + n5
            for matchNum, match in enumerate(Quantities, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip().replace(',', ''))  # Fix the "," problem
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n5 + 1, 5, float(val))
                n5 += 1

            n6 = n + n6
            for matchNum, match in enumerate(Varieties, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not '':
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n6 + 1, 6, val)
                    # w_sheet.write(n6 + 6, 5, match.group(groupNum))
                n6 += 1

            n7 = n + n7
            for matchNum, match in enumerate(KGs, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip().replace(',', ''))  # Fix the "," problem
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n7 + 1, 7, float(val))
                n7 += 1

            n8 = n + n8
            for matchNum, match in enumerate(Gradings, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n8 + 1, 8, val)
                n8 += 1

            n9 = n + n9
            for matchNum, match in enumerate(PFQ_scores, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n9 + 1, 9, float(val))
                n9 += 1

            n += 2

        # print('Process Success')
        # print(myPDF)

        #######################
        # Saving the Excel File

        wb.save(excelFile)
        pdfFileObj.close()

        status_update("The Excel file is saved at  :\n" + str(excelFile))

        show_processing_end()
        myPDF = None

        # Enable Open Excel File Button
        Button.config(text="  Select another PDF File  ")
        Button2.config(state="normal")
        Button2.config(command=lambda: open_excel(str(excelFile)))

        # open_excel(excelFile)

    except:
        if myPDF is not "":
            show_processing_error()


########################
# Creating a menu

menubar = Menu(window)

close_icon = tkinter.PhotoImage(file="img\\ic_close_black_16dp.png")
file_icon = tkinter.PhotoImage(file="img\\ic_insert_drive_file_black_18dp.png")
help_icon = tkinter.PhotoImage(file="img\\ic_help_16pt.png")


def myFacebook():
    url = 'http://www.facebook.com/mounirboulwafa'
    webbrowser.open_new(url)


filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label=" Select a PDF File ", image=file_icon, compound=tkinter.LEFT, command=load_pdf)
filemenu.add_separator()
filemenu.add_command(label=" Exit", image=close_icon, compound=tkinter.LEFT, command=window.quit)
menubar.add_cascade(label="File", menu=filemenu)

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label=" Help", image=help_icon, compound=tkinter.LEFT, command=myFacebook)
helpmenu.add_command(label=" Contact", underline=1, command=show_contact)
helpmenu.add_command(label=" About", command=show_about)
menubar.add_cascade(label="About", menu=helpmenu)

window.config(menu=menubar)


#######################
# Open Excel Application After end on processing


def open_excel(excelfile):
    xl = Dispatch("Excel.Application")
    xl.Visible = True  # otherwise excel is hidden

    # newest excel does not accept forward slash in path
    w = xl.Workbooks.Open(excelfile)


def status_update(mypdf):
    status_label.config(text=mypdf)
    # print("7777 " + mypdf)


Button = tkinter.Button(window, text="  Select a PDF File  ", command=load_pdf)
Button.pack()

space2 = tkinter.Label(window, height=1, ).pack()

status_label = Label(window, text="\n", )
status_label.pack(ipadx=80, ipady=16, fill='both')

Button2 = tkinter.Button(window, text="  Open the Excel File  ", state=DISABLED)
Button2.pack()

space1 = tkinter.Label(window, height=2).pack()

copyrights_label = tkinter.Label(window, text="Copyright © 2019 by Mounir Boulwafa. All rights reserved.",
                                 fg="white", bg="chartreuse4", width=300, height=1, font="Helvetica 8").pack()
# window.withdraw()
window.mainloop()
