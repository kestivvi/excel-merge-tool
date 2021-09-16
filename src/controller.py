from model import Model
from view_abc import View, BtnStatus
from controller_abc import Controller as Controller_abc

from workbook import Workbook

class Controller(Controller_abc):

    def __init__(self, model: Model, view: View):
        self.model = model
        self.view = view


    def handle_load_files_click(self):
        self.model.files = self.view.askForInputFiles()
        self.model.files = [f.name for f in self.model.files]
        self.view.updateFileList()
        self.view.checkBtnMergeStatus()
    

    def handle_choose_save_location(self):
        self.model.path_to_save = self.view.askForOutputFiles()
        self.view.setSaveLocationText(self.model.path_to_save)
        self.view.checkBtnMergeStatus()


    def handle_merge_click(self):        
        self.view.setBtnMergeStatus(BtnStatus.DISABLED)

        self.view.setProgressText("Loading first workbook...")
        self.model.output_workbook = Workbook(self.model.files[0])

        all_workbooks = len(self.model.files)
        def notify_about_progress(progress):
            text = f'Workbook: {current_workbook+2}/{all_workbooks} | Worksheet: {progress["sheet"]["current"]}/{progress["sheet"]["all"]} | Row: {progress["row"]["current"]}/{progress["row"]["all"]} | Column: {progress["column"]["current"]}/{progress["column"]["all"]}'
            self.view.setProgressText(text)


        # Copy data from each file to out excel file
        for current_workbook, workbook in enumerate(self.model.files[1:]):
            self.view.setProgressText(f"Loading workbook nr {current_workbook+2}...")
            Workbook.merge_workbook(self.model.output_workbook, Workbook(workbook), progress_callback_fn=notify_about_progress)

        # save new excel file
        self.view.setProgressText("Saving now...")
        self.model.output_workbook.save_workbook(save_path=self.model.path_to_save)
        self.view.setProgressText("Done!")
        self.view.notifySound()

        self.view.checkBtnMergeStatus()
    

    def start(self):
        self.view.setUpView(self)
        self.view.startMainLoop()
