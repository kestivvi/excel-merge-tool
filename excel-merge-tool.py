import tkinter, tkinter.filedialog, threading
import openpyxl as xl
from copy import copy, deepcopy
import os, sys 


### Global Variables
files = []
out_workbook = None
path_to_save = None
in_filetypes = (    ("Excel Workbook", "*.xlsx"),
                    ("Excel Macro-Enabled Workbook", "*.xlsm"),
                    ("Excel Macro-Enabled Workbook Template", "*.xltm"),
                    ("Excel Spreadsheet Template", "*.xltx"),
                    # ("OpenDocument Spreadsheet", "*.ods"),
                    # ("CSV (Comma delimited)", "*.csv"),
                    # ("Excel 97-2003 Workbook", "*.xls"),
                    # ("Excel Binary Workbook", "*.xlsb")
                    )

out_filetypes = (   ("Excel Workbook", "*.xlsx"),
                    # ("OpenDocument Spreadsheet", "*.ods"),
                    # ("CSV (Comma delimited)", "*.csv"),
                    )


### Functions' definitions
def load_files():
    global files
    files = tkinter.filedialog.askopenfiles(title="Select workbooks",filetypes=in_filetypes)
    
    for i in range(0, len(files)):
        files[i] = files[i].name

    if path_to_save is not None and files is not []:
        btn_merge['state'] = 'normal'
    
    list_files_box.delete(0, tkinter.END)
    for path in files:
        list_files_box.insert(tkinter.END, path)


def first_empty_row(ws):
    last = ws.max_row+2
    for row in range(1, last):
        if ws.cell(row=row, column=1).value is None:
            return row
    

def first_empty_column(ws):
    last = ws.max_column+2
    for column in range(1, last):
        if ws.cell(row=1, column=column).value is None:
            return column


def merge_files():
    length = len(files)
    wb_num = 1

    # Load first excel file
    progress_text.set(str(wb_num)+"/"+str(length)+" Workbook Loading...")
    global out_workbook
    out_workbook = xl.load_workbook(files[0])
    files.pop(0)


    # Copy data from each file to out excel file
    for workbook in files:
        wb_num += 1
        progress_text.set(str(wb_num)+"/"+str(length)+" Workbook Loading...")
        workbook = xl.load_workbook(workbook)
        
        ws_num = 1
        for sheet in workbook.worksheets:
            progress_text.set(str(wb_num)+"/"+str(length)+" Workbook   |   "+str(ws_num)+"/"+str(len(workbook.worksheets))+" Worksheet Copying...")
            ws_num += 1

            new_worksheet = sheet.title not in out_workbook.sheetnames
            if new_worksheet:
                out_workbook.create_sheet(sheet.title, workbook.sheetnames.index(sheet.title))
                start_row = 1
            else:
                start_row = 2
            
            outWorkSheet = out_workbook[sheet.title]
            workSheetRow = first_empty_row(outWorkSheet)-1
            last_row = first_empty_row(sheet)
            last_column = first_empty_column(sheet)


            for row in range(start_row, last_row):
                workSheetRow += 1

                for column in range(1, last_column):
                    cell = outWorkSheet.cell(row=workSheetRow, column=column)
                    cell.value = sheet.cell(row=row, column=column).value
                    
                    if new_worksheet:
                        previous_cell = sheet.cell(row=workSheetRow, column=column)
                    else:
                        previous_cell = outWorkSheet.cell(row=workSheetRow-2, column=column)
                        
                    cell.font = copy(previous_cell.font)
                    cell.border = copy(previous_cell.border)
                    cell.fill = copy(previous_cell.fill)
                    cell.number_format = copy(previous_cell.number_format)
                    cell.protection = copy(previous_cell.protection)
                    cell.alignment = copy(previous_cell.alignment)
    
    # save new excel file
    progress_text.set("Saving now...")
    out_workbook.save(filename=path_to_save)
    progress_text.set("Done!")
    window.bell()


def save_file():
    global path_to_save 
    path_to_save = tkinter.filedialog.asksaveasfilename(title="Save as", filetypes=out_filetypes, defaultextension='.xlsx')

    if path_to_save is not None and files is not []:
        btn_merge['state'] = 'normal'
    
    save_location_text.set(path_to_save)    


### Program starts
window = tkinter.Tk()
window.title("Excel Merge Tool")
window.iconbitmap("./icon.ico")

# Set GUI
btn_load = tkinter.Button(window, text="Select files", command=load_files, padx=15, pady=5)
btn_load.grid(row=0,column=0, padx=10, pady=10)

btn_save = tkinter.Button(window, text="Select save location", command=save_file, padx=15, pady=5)
btn_save.grid(row=0,column=1, padx=10, pady=10)

btn_merge = tkinter.Button(window, text="Merge and save", command=lambda : threading.Thread(target=merge_files).start(), padx=15, pady=5, state='disabled')
btn_merge.grid(row=0,column=2, padx=10, pady=10)

label1 = tkinter.Label(window, text="Selected files:")
label1.grid(row=1, column=0)

list_files_box = tkinter.Listbox(window, width=68)
list_files_box.grid(row=2,column=0, columnspan=3, sticky="W", padx=10)

label2 = tkinter.Label(window, text="Save location : ")
label2.grid(row=3, column=0, pady=5)

save_location_text = tkinter.StringVar(window)
label3 = tkinter.Label(window, textvariable=save_location_text)
label3.grid(row=3, column=1, sticky="W", pady=5, columnspan=2)
save_location_text.set("None")

progress_text = tkinter.StringVar(window)
progress_label = tkinter.Label(window, textvariable=progress_text, padx=10, pady=5)
progress_label.grid(row=4,column=0, columnspan=3, sticky='W')


# Program enters mainloop
window.mainloop()
