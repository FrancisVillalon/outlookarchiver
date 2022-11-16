import os
from tkinter import Tk

from AppController import AppController
from AppView import AppView
from UserInputModel import UserInput


# * Main
class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("Email Archive Tool")
        self.iconbitmap(os.path.join(os.getcwd(), "assets/email.ico"))

        view = AppView(self)
        view.grid(row=0, column=0, padx=10, pady=10)

        controller = AppController(UserInput, view)
        view.setController(controller)


if __name__ == "__main__":
    app = App()
    app.mainloop()
