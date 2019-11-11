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
excel_file = ""

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


def delete_worksheet_content(w_sheet):
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
    my_pdf = filedialog.askopenfilename(initialdir="C:\\Users\\" + str(os.getlogin()) + " \\Desktop",
                                       title=" Select file",
                                       filetypes=(("PDF Documents", "*.pdf"), ("all files", "*.*")))

    # print("File is loaded")

    #######################
    # Read the PDF Document

    try:
        pdf_file_obj = open(my_pdf, 'rb')
        # pdfFileObj = open('pdf.pdf', 'rb')
        pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj)
        pages = pdf_reader.numPages

        #######################
        # Save the Data As Excel File

        excel_file = str(my_pdf).replace('.pdf', '.xls')
        # print("1111" + str(excel_file))

        #######################
        # Create the Excel File

        # if the Excel file exist : update
        if path.exists(excel_file):
            rb = xlrd.open_workbook(excel_file, formatting_info=1)
            wb = copy(rb)
            w_sheet = wb.get_sheet(0)

            delete_worksheet_content(w_sheet)
            # print("File Updated")

        # if the Excel file doesn't exist : create new one
        else:
            wb = Workbook()
            w_sheet = wb.add_sheet('Daily Quality Report')

            write_headers(w_sheet)

            wb.save(excel_file)

        ############################################################################
        # start processing

        # Inserting the logo
        # w_sheet.insert_bitmap('img\\logo.bmp', 0, 0, 2, 2)

        n1 = 0
        n2 = 0
        n3 = 0
        n4 = 0
        n5 = 0
        n6 = 0
        n7 = 0
        n8 = 0
        n9 = 0
        n10 = 0

        # Reading Grower & Ranche

        grower_val = NONE
        ranch_val = NONE

        page_obj = pdf_reader.getPage(0)
        page_one = page_obj.extractText()
        # print(pageOne)

        grower_regex = r"Grower:.*\n(\w+)|Producteur:.*\n(\w+)"
        growers = re.finditer(grower_regex, page_one)

        ranch_regex = r"Ranch:.*\n(\w+)|Ferme:.*\n(\w+)"
        ranches = re.finditer(ranch_regex, page_one)

        for match_num, match in enumerate(growers, start=1):
            for group_num in range(0, len(match.groups())):
                group_num = group_num + 1
                if match.group(group_num) is not None:
                    grower_val = int(match.group(group_num))

        for match_num, match in enumerate(ranches, start=1):
            for group_num in range(0, len(match.groups())):
                group_num = group_num + 1
                if match.group(group_num) is not None:
                    ranch_val = int(match.group(group_num))

        # print(Grower_val)
        # print(Ranche_val)

        for x in range(0, pages):
            # print("\n--------- Page " + str(x + 1) + " ----------")
            page_obj = pdf_reader.getPage(x)
            text = page_obj.extractText()
            # print(text)

            ######################
            #  RegExpressions

            grower_receipt_regex = r"Grower receipt:.*\n(.*)|Bon de réception:.*\n(.*)"
            grower_receipts = re.finditer(grower_receipt_regex, text)

            item_number_regex = r"Production method.*\n.*.*\n(.*)|Méthode de Production.*\n.*.*\n(.*)"
            item_numbers = re.finditer(item_number_regex, text)

            id_bloc_regex = r"Batch number:.*\n.*(.{8}).{2}|Numéro de Lot:.*\n.*(.{8}).{2}"
            id_blocs = re.finditer(id_bloc_regex, text)

            batch_number_regex = r"Batch number:.*\n(.*)|Numéro de Lot:.*\n(.*)"
            batch_numbers = re.finditer(batch_number_regex, text)

            qc_regex = r"QC check:.*\n(.*)|Contrôle qualité:.*\n(.*)"
            qcs = re.finditer(qc_regex, text)

            quantity_regex = r"(.*\n.*)MA MOU|(.*\n.*)MA DAC|(.*\n.*)MA LAR"
            quantities = re.finditer(quantity_regex, text)

            variety_regex = r"MA MOU.*\n(.*)|MA DAC.*\n(.*)|MA LAR.*\n(.*)"
            varieties = re.finditer(variety_regex, text)

            kg_regex = r"Quantity in KGs:.*\n(.*)|Quantité en KG:.*\n(.*)"
            kgs = re.finditer(kg_regex, text)

            grading_regex = r"Final Grading:.*\n(.*)|Classification  finale:.*\n(.*)"
            gradings = re.finditer(grading_regex, text)

            pfq_score_regex = r"Final PFQ score:.*\n(.*)|Score PFQ final:.*\n(.*)"
            pfq_scores = re.finditer(pfq_score_regex, text)

            n = 0
            n1 = n + n1
            for match_num, match in enumerate(grower_receipts, start=1):
                for group_num in range(0, len(match.groups())):
                    group_num = group_num + 1
                    # print(match.group(group_num))
                    if match.group(group_num) is not None:
                        val = str(match.group(group_num).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n1 + 1, 0, val)

                            w_sheet.write(n1 + 1, 10, grower_val)
                            w_sheet.write(n1 + 1, 11, ranch_val)

                n1 += 1

            n2 = n + n2
            for match_num, match in enumerate(qcs, start=1):
                for group_num in range(0, len(match.groups())):
                    group_num = group_num + 1
                    # print(match.group(group_num))
                    if match.group(group_num) is not None:
                        val = match.group(group_num).rstrip()
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n2 + 1, 4, val)
                n2 += 1

            n10 = n + n10
            for match_num, match in enumerate(batch_numbers, start=1):
                # print("111111111")
                for group_num in range(0, len(match.groups())):
                    group_num = group_num + 1
                    # print(match.group(group_num))
                    if match.group(group_num) is not None:
                        val = str(match.group(group_num).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n10 + 1, 3, val)
                n10 += 1

            n3 = n + n3
            for match_num, match in enumerate(id_blocs, start=1):
                # print("111111111")
                for group_num in range(0, len(match.groups())):
                    group_num = group_num + 1
                    # print(match.group(group_num))
                    if match.group(group_num) is not None:
                        val = str(match.group(group_num).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n3 + 1, 2, val)
                n3 += 1

            n4 = n + n4
            for match_num, match in enumerate(item_numbers, start=1):
                for group_num in range(0, len(match.groups())):
                    group_num = group_num + 1
                    # print(match.group(group_num))
                    if match.group(group_num) is not None:
                        val = str(match.group(group_num).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n4 + 1, 1, int(val))
                n4 += 1

            n5 = n + n5
            for match_num, match in enumerate(quantities, start=1):
                for group_num in range(0, len(match.groups())):
                    group_num = group_num + 1
                    # print(match.group(group_num))
                    if match.group(group_num) is not None:
                        val = str(match.group(group_num).rstrip().replace(',', ''))  # Fix the "," problem
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n5 + 1, 5, float(val))
                n5 += 1

            n6 = n + n6
            for match_num, match in enumerate(varieties, start=1):
                for group_num in range(0, len(match.groups())):
                    group_num = group_num + 1
                    # print(match.group(group_num))
                    if match.group(group_num) is not None:
                        val = str(match.group(group_num).rstrip())
                        if val is not '':
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n6 + 1, 6, val)
                    # w_sheet.write(n6 + 6, 5, match.group(group_num))
                n6 += 1

            n7 = n + n7
            for match_num, match in enumerate(kgs, start=1):
                for group_num in range(0, len(match.groups())):
                    group_num = group_num + 1
                    # print(match.group(group_num))
                    if match.group(group_num) is not None:
                        val = str(match.group(group_num).rstrip().replace(',', ''))  # Fix the "," problem
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n7 + 1, 7, float(val))
                n7 += 1

            n8 = n + n8
            for match_num, match in enumerate(gradings, start=1):
                for group_num in range(0, len(match.groups())):
                    group_num = group_num + 1
                    # print(match.group(group_num))
                    if match.group(group_num) is not None:
                        val = str(match.group(group_num).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n8 + 1, 8, val)
                n8 += 1

            n9 = n + n9
            for match_num, match in enumerate(pfq_scores, start=1):
                for group_num in range(0, len(match.groups())):
                    group_num = group_num + 1
                    # print(match.group(group_num))
                    if match.group(group_num) is not None:
                        val = str(match.group(group_num).rstrip())
                        if val is not None:
                            # print(str("---") + val + str(" 00"))
                            w_sheet.write(n9 + 1, 9, float(val))
                n9 += 1

            n += 2

        # print('Process Success')
        # print(myPDF)

        #######################
        # Saving the Excel File

        wb.save(excel_file)
        pdf_file_obj.close()

        status_update("The Excel file is saved at  :\n" + str(excel_file))

        show_processing_end()
        my_pdf = None

        # Enable Open Excel File Button
        Button.config(text="  Select another PDF File  ")
        Button2.config(state="normal")
        Button2.config(command=lambda: open_excel(str(excel_file)))

        # open_excel(excel_file)

    except:
        if myPDF is not "":
            show_processing_error()


########################
# Creating a menu

menu_bar = Menu(window)

close_icon = tkinter.PhotoImage(file="img\\ic_close_black_16dp.png")
file_icon = tkinter.PhotoImage(file="img\\ic_insert_drive_file_black_18dp.png")
help_icon = tkinter.PhotoImage(file="img\\ic_help_16pt.png")


def my_facebook():
    url = 'http://www.facebook.com/mounirboulwafa'
    webbrowser.open_new(url)


file_menu = Menu(menu_bar, tearoff=0)
file_menu.add_command(label=" Select a PDF File ", image=file_icon, compound=tkinter.LEFT, command=load_pdf)
file_menu.add_separator()
file_menu.add_command(label=" Exit", image=close_icon, compound=tkinter.LEFT, command=window.quit)
menu_bar.add_cascade(label="File", menu=file_menu)

help_menu = Menu(menu_bar, tearoff=0)
help_menu.add_command(label=" Help", image=help_icon, compound=tkinter.LEFT, command=my_facebook)
help_menu.add_command(label=" Contact", underline=1, command=show_contact)
help_menu.add_command(label=" About", command=show_about)
menu_bar.add_cascade(label="About", menu=help_menu)

window.config(menu=menu_bar)


#######################
# Open Excel Application After end on processing


def open_excel(excel_file):
    xl = Dispatch("Excel.Application")
    xl.Visible = True  # otherwise excel is hidden

    # newest excel does not accept forward slash in path
    w = xl.Workbooks.Open(excel_file)


def status_update(my_pdf):
    status_label.config(text=my_pdf)
    # print("7777 " + my_pdf)


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
