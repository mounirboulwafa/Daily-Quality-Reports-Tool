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

#######################
# Open file Dialog to choose the PDF Document

window = tkinter.Tk()
window.title(" ATLAS BERRY FARMS | Daily Quality Reports Tool")
window.geometry("500x400")
window.iconbitmap(default='img\\icon2.ico')
window.resizable(0, 0)


#######################
# Open about window

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


def show_about():
    win = Toplevel(window)
    win.title(" About")
    win.geometry("400x200")
    win.resizable(0, 0)

    title = tkinter.Label(win, text="ATLAS BERRY FARMS\nDaily Quality Reports Tool v1.0",
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
    messagebox.showwarning(" Error", " Process could not be completed !\n"
                                     " The Excel file is already open ! ")


logo = tkinter.PhotoImage(file="img\\logo.png")

logo_label = tkinter.Label(window, image=logo, width=500, height=110, )
logo_label.place(x=20, y=60)
logo_label.pack()

title_label = tkinter.Label(window, text="Daily Quality Reports Tool",
                            fg="white", bg="chartreuse4", width=300, height=2, font="Helvetica 16 bold").pack()

space = tkinter.Label(window, height=2, ).pack()

myPDF = None
excelFile = ""


def delete_worksheet(w_sheet):
    w_sheet.write(3, 3, "")
    w_sheet.write(4, 3, "")
    for x in range(0, 9):
        for y in range(6, 60):
            w_sheet.write(y, x, "")


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

            style1 = xlwt.easyxf(
                'pattern: pattern solid, fore_colour green;'
                'font: colour white,height 260, bold True;'
                'align: horiz center')

            style = xlwt.easyxf(
                'pattern: pattern solid, fore_colour green;'
                'font: colour white,height 220, bold True;'
                'align: horiz center')

            #######################
            # Writing titles

            # w_sheet.write(0, 3, 'Daily Quality Report', style)
            w_sheet.write_merge(0, 0, 3, 4, 'Daily Quality Report', style1)
            w_sheet.write(2, 3, 'Grower', style)
            w_sheet.write(2, 4, 'Ranch', style)
            w_sheet.write(5, 0, 'Grower receipt', style)
            w_sheet.write(5, 1, 'Item number', style)
            w_sheet.write(5, 2, 'Batch number', style)
            w_sheet.write(5, 3, 'Arrival date / QC check', style)
            w_sheet.write(5, 4, 'Quantity', style)
            w_sheet.write(5, 5, 'Variety', style)
            w_sheet.write(5, 6, 'Quantity in KGs', style)
            w_sheet.write(5, 7, 'Final Grading', style)
            w_sheet.write(5, 8, 'Final PFQ score', style)

            #######################
            # Styling ( Column width )

            w_sheet.col(0).width = 5000
            w_sheet.col(1).width = 4000
            w_sheet.col(2).width = 4000
            w_sheet.col(3).width = 6400
            w_sheet.col(4).width = 4000
            w_sheet.col(5).width = 4000
            w_sheet.col(6).width = 4800
            w_sheet.col(7).width = 4500
            w_sheet.col(8).width = 4800

            wb.save(excelFile)

        ###################
        # start processing

        # Inserting the logo
        w_sheet.insert_bitmap('img\\logo.bmp', 0, 0, 2, 2)

        n1 = 0
        n2 = 0
        n3 = 0
        n4 = 0
        n5 = 0
        n6 = 0
        n7 = 0
        n8 = 0
        n9 = 0

        for x in range(0, pages):
            # print("\n--------- Page " + str(x + 1) + " ----------")
            page_obj = pdfReader.getPage(x)
            text = page_obj.extractText()
            # print(text)

            ######################
            #  RegExpressions

            grower_regex = r"Grower:.*\n(\w+)|Producteur:.*\n(\w+)"
            ranch_regex = r"Ranch:.*\n(\w+)|Ferme:.*\n(\w+)"

            grower_receipt_regex = r"Grower receipt:.*\n(.*)|Bon de réception:.*\n(.*)"
            qc_regex = r"QC check:.*\n(.*)|Contrôle qualité:.*\n(.*)"
            batch_number_regex = r"Batch number:.*\n.*(.{8}).{2}|Numéro de Lot:.*\n.*(.{8}).{2}"

            item_number_regex = r"Production method.*\n.*.*\n(.*)|Méthode de Production.*\n.*.*\n(.*)"
            quantity_regex = r"(.*\n.*)MA MOU|(.*\n.*)MA DAC"
            variety_regex = r"MA MOU.*\n(.*)|MA DAC.*\n(.*)"

            kg_regex = r"Quantity in KGs:.*\n(.*)|Quantité en KG:.*\n(.*)"
            grading_regex = r"Final Grading:.*\n(.*)|Classification  finale:.*\n(.*)"
            pfq_score_regex = r"Final PFQ score:.*\n(.*)|Score PFQ final:.*\n(.*)"

            ######################
            #

            growers = re.finditer(grower_regex, text)
            ranches = re.finditer(ranch_regex, text)

            grower_receipts = re.finditer(grower_receipt_regex, text)
            qcs = re.finditer(qc_regex, text)
            batch_numbers = re.finditer(batch_number_regex, text)
            # print(BatchNumbers)

            item_numbers = re.finditer(item_number_regex, text)
            quantities = re.finditer(quantity_regex, text)
            varieties = re.finditer(variety_regex, text)

            kgs = re.finditer(kg_regex, text)
            gradings = re.finditer(grading_regex, text)
            pfq_scores = re.finditer(pfq_score_regex, text)

            ######################
            #

            n = 0
            for matchNum, match in enumerate(growers, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    # w_sheet.write(2, 4, int(match.group(groupNum)))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(3, 3, int(val))

            for matchNum, match in enumerate(ranches, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    # w_sheet.write(3, 4, int(match.group(groupNum)))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(3, 4, int(val))

            n1 = n + n1
            for matchNum, match in enumerate(grower_receipts, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n1 + 6, 0, val)
                n1 += 1

            n2 = n + n2
            for matchNum, match in enumerate(qcs, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n2 + 6, 3, val)
                n2 += 1

            n3 = n + n3
            for matchNum, match in enumerate(batch_numbers, start=1):
                # print("111111111")
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n3 + 6, 2, val)
                n3 += 1

            n4 = n + n4
            for matchNum, match in enumerate(item_numbers, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n4 + 6, 1, int(val))
                n4 += 1

            n5 = n + n5
            for matchNum, match in enumerate(quantities, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n5 + 6, 4, float(val))
                n5 += 1

            n6 = n + n6
            for matchNum, match in enumerate(varieties, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not '':
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n6 + 6, 5, val)
                    # w_sheet.write(n6 + 6, 5, match.group(groupNum))
                n6 += 1

            n7 = n + n7
            for matchNum, match in enumerate(kgs, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip().replace(',', ''))
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n7 + 6, 6, float(val))
                n7 += 1

            n8 = n + n8
            for matchNum, match in enumerate(gradings, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n8 + 6, 7, val)
                n8 += 1

            n9 = n + n9
            for matchNum, match in enumerate(pfq_scores, start=1):
                for groupNum in range(0, len(match.groups())):
                    groupNum = groupNum + 1
                    # print(match.group(groupNum))
                    if match.group(groupNum) is not None:
                        val = str(match.group(groupNum).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n9 + 6, 8, float(val))
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
