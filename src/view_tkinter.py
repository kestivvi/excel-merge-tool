from typing import List
import tkinter, tkinter.filedialog
import threading

from view_abc import View as View_abc, BtnStatus
from model import Model
from controller_abc import Controller

class View(View_abc):
    # model: Model
    # window: tkinter.Tk
    # btn_merge: tkinter.Button
    # progress_text: tkinter.StringVar
    # save_location_text = tkinter.StringVar


    def __init__(self, model: Model, title: str = "Excel Merge Tool", icon_path: str = "./icon.ico"):
        self.model = model
        self.window = tkinter.Tk()

        self.window.title(title)
        self.window.iconbitmap(icon_path)

        self.bool_copy_style = tkinter.IntVar()
        self.bool_copy_style.set(1)


    def setUpView(self, controller: Controller):
        btn_load = tkinter.Button(self.window, text="Select files", command=controller.handle_load_files_click, padx=15, pady=5)
        btn_load.grid(row=0,column=0, padx=10, pady=10)

        btn_save = tkinter.Button(self.window, text="Select save location", command=controller.handle_choose_save_location, padx=15, pady=5)
        btn_save.grid(row=0,column=1, padx=10, pady=10)

        self.btn_merge = tkinter.Button(self.window, text="Merge and save", command=lambda : threading.Thread(target=controller.handle_merge_click).start(), padx=15, pady=5, state='disabled')
        self.btn_merge.grid(row=0,column=2, padx=10, pady=10)

        label1 = tkinter.Label(self.window, text="Selected files:")
        label1.grid(row=1, column=0)

        self.list_files_box = tkinter.Listbox(self.window, width=68)
        self.list_files_box.grid(row=2,column=0, columnspan=3, sticky="W", padx=10)

        label2 = tkinter.Label(self.window, text="Save location : ")
        label2.grid(row=3, column=0, pady=5)

        self.save_location_text = tkinter.StringVar(self.window)
        label3 = tkinter.Label(self.window, textvariable=self.save_location_text)
        label3.grid(row=3, column=1, sticky="W", pady=5, columnspan=2)
        self.save_location_text.set("None")

        checkbox_copy_style = tkinter.Checkbutton(self.window, variable=self.bool_copy_style, onvalue=1, offvalue=0)
        checkbox_copy_style.grid(row=4, column=1, sticky='W')

        label4 = tkinter.Label(self.window, text="Copy cells' style")
        label4.grid(row=4, column=0, pady=5, columnspan=1)

        self.progress_text = tkinter.StringVar(self.window)
        progress_label = tkinter.Label(self.window, textvariable=self.progress_text, padx=10, pady=5)
        progress_label.grid(row=5,column=0, columnspan=3, sticky='W')


    def get_copy_style_bool(self):
        return self.bool_copy_style.get()


    def askForInputFiles(self, title: str = "Select workbooks") -> List[str]:
        return tkinter.filedialog.askopenfiles(title=title, filetypes=self.model.in_filetypes)


    def askForOutputFiles(self, title: str = "Save as") -> List[str]:
        return tkinter.filedialog.asksaveasfilename(title=title, filetypes=self.model.out_filetypes, defaultextension='.xlsx')
    

    def setBtnMergeStatus(self, status: BtnStatus):
        self.btn_merge["state"] = status.value


    def checkBtnMergeStatus(self):
        if self.model.path_to_save is not None and self.model.files is not []:
            self.setBtnMergeStatus(BtnStatus.NORMAL)


    def updateFileList(self):
        self.list_files_box.delete(0, tkinter.END)
        [self.list_files_box.insert(tkinter.END, file) for file in self.model.files]


    def setProgressText(self, text: str):
        self.progress_text.set(text)
    

    def setSaveLocationText(self, text: str):
        self.save_location_text.set(text)


    def notifySound(self):
        self.window.bell()


    def startMainLoop(self):
        self.window.mainloop()
