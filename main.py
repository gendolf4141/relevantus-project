from customtkinter import CTk, CTkButton
import tkinter.filedialog as fd
import reports
from service import check_busy_file
from pathlib import Path
from enum import Enum


class TitleButton(str, Enum):
    CREATE = "Сформировать"
    LOAD = "Скачать"
    ANALYSIS = "Анализировать"


class App(CTk):
    def __init__(self):
        super().__init__()
        self.geometry("400x150")
        self.title("RELEVANTUS")

        self.columnconfigure(0, weight=1)

        self.button = CTkButton(
            self,
            text=TitleButton.CREATE,
            command=self.create
        )
        self.button.grid(row=0, column=0, padx=20, pady=10, sticky="ew")

        self.button = CTkButton(
            self,
            text=TitleButton.LOAD,
            command=self.load
        )
        self.button.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        self.button = CTkButton(
            self,
            text=TitleButton.ANALYSIS,
            command=self.analysis
        )
        self.button.grid(row=3, column=0, padx=20, pady=10, sticky="ew")

    @staticmethod
    def select_file_excel() -> Path:
        file = fd.askopenfilename(
            title="Выберите файл",
            filetypes=[("excel", "*.xlsx")],
        )
        return Path(file)

    def create(self) -> None:
        file = self.select_file_excel()
        check_busy_file(file)
        reports.create(file)

    def load(self) -> None:
        file = self.select_file_excel()
        check_busy_file(file)
        reports.load(file)

    def analysis(self) -> None:
        file = self.select_file_excel()
        check_busy_file(file)
        reports.analysis(file)


if __name__ == "__main__":
    app = App()
    app.mainloop()