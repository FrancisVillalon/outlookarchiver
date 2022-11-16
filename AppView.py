import os
from tkinter import *
from tkinter import Text, filedialog, messagebox, ttk


# * View
class AppView(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        # * vars
        self.folderPath = StringVar()
        self.smBool = BooleanVar()
        self.emailCat = StringVar()
        self.emailCategoryList = ["Trade", "Non-Trade", "CA"]

        # * default vals
        self.emailCat.set(self.emailCategoryList[0])
        self.folderPath.set(os.path.join(os.getcwd(), "archived-mail"))

        # export path browser
        self.exportFolder_browser = Button(
            self, text="Browse Folder", command=self.getFolderPath
        )
        self.exportFolder_browser.grid(row=0, column=2, sticky="w", padx=5, pady=5)

        # export path entry field
        self.exportFolder_entry_label = Label(self, text="Export Folder: ")
        self.exportFolder_entry_label.grid(row=0, column=0, sticky="w")
        self.exportFolder_entry = Entry(self, textvariable=self.folderPath)
        self.exportFolder_entry.grid(row=0, column=1, ipadx=100)

        # progress bar
        self.pb = ttk.Progressbar(
            self, orient=HORIZONTAL, mode="determinate", length=100
        )
        self.pb.grid(row=3, column=1, columnspan=3, sticky="ew", padx=(0, 5))
        self.pb_label = Label(self, text="Task Progress: ").grid(
            row=3, column=0, sticky="w"
        )

        # supress messages
        self.supress_msgs_label = Label(self, text="Options: ").grid(
            row=4, column=0, sticky="w"
        )
        self.supress_msgs = Checkbutton(
            self, text="Supress messages", variable=self.smBool
        ).grid(row=4, column=1, sticky="w")

        # custom folder name
        self.CustomFname = Entry(self)
        self.CustomFname.grid(row=5, column=1, pady=5, sticky="w")
        self.CustomFname_label = Label(self, text="Create New Folder: ").grid(
            row=5, column=0, sticky="w"
        )

        # bpnum
        self.bpNum = Entry(self)
        self.bpNum.grid(row=6, column=1, pady=5, sticky="w")
        self.bpNum_label = Label(self, text="Enter BP Num: ").grid(
            row=6, column=0, sticky="w"
        )

        # ordernum
        self.orderNum = Entry(self)
        self.orderNum.grid(row=7, column=1, pady=5, sticky="w")
        self.orderNum_label = Label(self, text="Enter Order Num: ").grid(
            row=7, column=0, sticky="w"
        )

        # emailcat
        self.emailCat_menu = OptionMenu(self, self.emailCat, *self.emailCategoryList)
        self.emailCat_label = Label(self, text="Enter Email Category: ").grid(
            row=8, column=0, sticky="w"
        )
        self.emailCat_menu.grid(row=8, column=1, pady=5, sticky="w")

        # execute btn
        self.executeBtn = Button(self, text="Export", command=self.exportEmails_clicked)
        self.executeBtn.grid(row=9, column=0, pady=5, sticky="w")

        # controller
        self.controller = None

    def setController(self, controller):
        self.controller = controller

    def getFolderPath(self):
        folder_selected = filedialog.askdirectory()
        self.folderPath.set(folder_selected)

    def exportEmails_clicked(self):
        if self.controller:
            self.controller.exportEmails(
                self.folderPath.get(),
                self.smBool.get(),
                self.CustomFname.get(),
                self.bpNum.get(),
                self.orderNum.get(),
                self.emailCat.get(),
            )

    def showError(self, msg):
        messagebox.showerror("Error", msg)

    def showInfo(self, msg):
        messagebox.showinfo("Info", msg)

    def showConfirm(self, msg):
        ans = messagebox.askyesno("Confirm", msg)
        return ans
