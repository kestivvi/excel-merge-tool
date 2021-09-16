from abc import ABC, abstractmethod
from enum import Enum
from typing import List

class BtnStatus(Enum):
    DISABLED = "disabled"
    NORMAL = "normal"

class View(ABC):

    @abstractmethod
    def __init__(self, title: str = "Excel Merge Tool", icon_path: str = "./icon.ico"):
        """Initialize window, set up title and icon."""


    @abstractmethod
    def setUpView(self, controller):
        """Set up gui objects."""


    @abstractmethod
    def startMainLoop(self):
        """Start the main loop of the program."""


    @abstractmethod
    def get_copy_style_bool(self):
        """Get the value of a checkbox are we gonna copy cell's style."""


    @abstractmethod
    def askForInputFiles(self, title: str = "Select workbooks") -> List[str]:
        """Ask user to select input files."""


    @abstractmethod    
    def askForOutputFiles(self, title: str = "Save as") -> List[str]:
        """Ask user to choose where to save output file."""


    @abstractmethod
    def checkBtnMergeStatus(self):
        """Check and set the status of the merge button."""
    

    @abstractmethod
    def setBtnMergeStatus(self, status: BtnStatus):
        """Set the status of the merge button."""


    @abstractmethod
    def updateFileList(self):
        """Update list with input files."""


    @abstractmethod
    def setProgressText(self, text: str):
        """Set the text displayed as the progress."""


    @abstractmethod
    def setSaveLocationText(self, text: str):
        """Set the text with the choosen location of output file."""
    

    @abstractmethod
    def notifySound(self):
        """Play the notification sound."""
