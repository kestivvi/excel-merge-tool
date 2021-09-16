from abc import ABC, abstractmethod

from model import Model
from view_abc import View

class Controller(ABC):

    @abstractmethod
    def __init__(self, model: Model, view: View) -> None:
        """Initialize controller."""

    def handle_load_files_click(self) -> None:
        """Update model with selected files by the user."""

    def handle_choose_save_location(self) -> None:
        """Update model with save location selected by the user."""

    def handle_merge_click(self) -> None:
        """Merge files selected by the user and save new file at given location."""

    def start(self) -> None:
        """Start program."""
