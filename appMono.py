import os
import string
from tkinter import *
from tkinter import Text, Tk, filedialog, messagebox, ttk

import pandas as pd
import win32com.client

from AppController import AppController
from AppLogger import logger
from AppView import AppView
from UserInputModel import UserInput


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


# * Model
class UserInput:
    def __init__(self, exportFolder, bpNum, orderNum, emailCat):
        self.exportFolder = exportFolder
        self.bpNum = bpNum
        self.orderNum = orderNum
        self.emailCat = emailCat

    @property
    def exportFolder(self):
        return self.__exportFolder

    @property
    def bpNum(self):
        return self.__bpNum

    @property
    def orderNum(self):
        return self.__orderNum

    @property
    def emailCat(self):
        return self.__emailCat

    @exportFolder.setter
    def exportFolder(self, value):
        if len(value) == 0:
            raise ValueError(f"Please define Export Folder")
        elif os.path.exists(value):
            self.__exportFolder = value
        else:
            raise ValueError(f"Export Folder does not exist.")

    @bpNum.setter
    def bpNum(self, value):
        self.__bpNum = value

    @orderNum.setter
    def orderNum(self, value):
        self.__orderNum = value

    @emailCat.setter
    def emailCat(self, value):
        self.__emailCat = value


# * Controller
class AppController:
    def __init__(self, UserInput, AppView):
        self.model = UserInput
        self.view = AppView
        self.obj = win32com.client.Dispatch("Outlook.Application")
        self.objWord = win32com.client.Dispatch("Word.Application")
        self.outlook = self.obj.GetNamespace("MAPI")

    def exportEmails(self, folderPath, smBool, CustomFname, bpNum, orderNum, emailCat):
        try:

            # * check for default folder path
            defaultFolderPath = os.path.join(os.getcwd(), "archived-mail")
            if folderPath == defaultFolderPath:
                if not os.path.exists(defaultFolderPath):
                    createDefaultPrompt = self.view.showConfirm(
                        "'archived-mail' folder does not exist in current directory. \n\n Would you like to create an 'archived-mail' folder in the current directory?"
                    )
                    if createDefaultPrompt:
                        os.mkdir(defaultFolderPath)
                    else:
                        folderPath = os.getcwd()
                        self.view.folderPath.set(folderPath)

            # * vars
            model = self.model(folderPath, bpNum, orderNum, emailCat)
            selectedMail = self.obj.ActiveExplorer().Selection
            numEmails = len(list(selectedMail))
            pb_d = 100 // numEmails

            # * Define master list
            masterListColumns = [
                "OrderNumber",
                "AccountNumber",
                "Category",
                "SenderName",
                "EmailSubject",
                "ReceiveDate",
                "Email Archive No.",
                "Attachments No.",
                "Files Attached",
                "SavePath",
                "EntryID",
            ]
            atchListColumns = ["Email Archive No.", "Attachment No.", "SavePath"]
            masterListPath = os.path.join(folderPath, "masterList.xlsx")
            if "masterList.xlsx" not in os.listdir(folderPath):
                masterListDf = pd.DataFrame(None, columns=masterListColumns)
                atchListDf = pd.DataFrame(None, columns=atchListColumns)
            else:
                masterListDf = pd.read_excel(
                    masterListPath, sheet_name="Sheet1", engine="openpyxl"
                )
                atchListDf = pd.read_excel(
                    os.path.join(masterListPath), sheet_name="Sheet2", engine="openpyxl"
                )
            if masterListDf.empty:
                offsetidx = 0
            else:
                offsetidx = masterListDf["Email Archive No."].values[-1]

            # * Check for selected emails
            if not numEmails > 0:
                self.view.showError("Error", "No emails selected in outlook.")

            # * Check for supress message
            if smBool:
                ans = True
            else:
                ans = self.view.showConfirm(
                    f"Do you want to archive {numEmails} selected emails?"
                )

            # * start archive operation
            if ans:

                # * Handling custom f name
                if (
                    len(
                        CustomFname.translate(str.maketrans("", "", string.punctuation))
                    )
                    > 0
                ):
                    folderPath = os.path.join(model.exportFolder, CustomFname)
                    if not os.path.exists(folderPath):
                        os.mkdir(folderPath)
                else:
                    folderPath = model.exportFolder

                # * Handling children folders
                for childName in self.view.emailCategoryList:
                    childPath = os.path.join(folderPath, childName)
                    if not os.path.exists(childPath):
                        os.mkdir(childPath)

                # * Child folderpath
                folderPathChild = os.path.join(folderPath, emailCat)

                # * Email
                for idx, email in enumerate(selectedMail):
                    try:
                        # * email data
                        emailId = email.EntryID
                        emailSender = f"{email.SenderName} ({email.SenderEmailAddress})"
                        emailGroup = idx + offsetidx + 1
                        emailSubject = email.subject
                        emailReceive = email.ReceivedTime.strftime("%b-%d-%Y %H:%M:%S")
                        emailRecieveFileName = email.ReceivedTime.strftime(
                            "%b%d%Y-%H%M%S"
                        )
                        emailFullFileName = (
                            f"{model.bpNum}_{model.orderNum}_{model.emailCat}_{emailGroup}_"
                            + emailRecieveFileName
                        )
                        emailPathMsg = os.path.join(folderPathChild, emailFullFileName)

                        # * archive email
                        if not masterListDf.loc[
                            masterListDf["EntryID"] == emailId
                        ].empty:
                            logger.info(
                                f"Archive request for {emailId} denied because it already has been archived. Refer to masterlist."
                            )
                            continue
                        else:
                            masterListDf.loc[masterListDf.shape[0]] = [
                                orderNum,
                                bpNum,
                                emailCat,
                                emailSender,
                                emailSubject,
                                emailReceive,
                                emailGroup,
                                len(list(email.Attachments)),
                                (
                                    " , ".join(
                                        [atch.FileName for atch in email.Attachments]
                                    )
                                ),
                                emailPathMsg.replace("/", "\\"),
                                emailId,
                            ]
                        email.SaveAs(emailPathMsg + ".msg", 3)
                        logger.info(
                            f"""
                            Archived email with parameters: \n
                            emailId= {emailId}\n
                            orderNum= {orderNum}, \n
                            bpNum= {bpNum}, \n
                            emailCat= {emailCat}, \n 
                            emailSender= {emailSender}, \n 
                            emailSubject= {emailSubject}, \n
                            emailReceive= {emailReceive}, \n
                            emailGroup = {emailGroup} \n
                            """
                        )
                        # * Convert msg to pdf
                        email.SaveAs(emailPathMsg + ".mht", 10)
                        self.MsgToPDF(
                            emailPathMsg + ".mht",
                            folderPathChild,
                            emailFullFileName + ".pdf",
                        )

                        # * Attachments
                        for idx, atch in enumerate(email.Attachments):
                            atchPathMsg = os.path.join(
                                folderPathChild,
                                f"{model.bpNum}_{model.orderNum}_{model.emailCat}_{emailGroup}_A{idx+1}_{atch.FileName}",
                            )
                            atch.SaveASFile(atchPathMsg)
                            atchListDf.loc[atchListDf.shape[0]] = [
                                emailGroup,
                                f"A{idx+1}",
                                atchPathMsg.replace("/", "\\"),
                            ]
                        # * update progress bar
                        if (idx + 1) != numEmails or numEmails == 1:
                            self.view.pb["value"] = pb_d * (idx + 1)
                        else:
                            self.view.pb["value"] = 100
                        self.view.update_idletasks()
                    except Exception as e:
                        logger.warning(
                            f"Failed to save email {idx+offsetidx} due to Exception =  {e} @ {e.__traceback__.tb_lineno}"
                        )
                        continue
                # * Update and save masterlist
                with pd.ExcelWriter(masterListPath, engine="openpyxl") as writer:
                    masterListDf.to_excel(writer, index=False, sheet_name="Sheet1")
                    atchListDf.to_excel(writer, index=False, sheet_name="Sheet2")
                    logger.info(f"Masterlist updated and saved @ {masterListPath} ")
                # ? Success
                logger.info(
                    f"Archive operation successful for {numEmails} emails @ {folderPath}"
                )
                if not smBool:
                    self.view.showInfo(
                        f"{numEmails} successfully archived at directory: \n\n {folderPath}"
                    )
        except ValueError as error:
            logger.warning(
                f"Failed to perform operation due to Exception =  {e} @ {e.__traceback__.tb_lineno}"
            )
            self.view.showError(
                "Operation has failed due to an unknown error. Check log for more details."
            )

    def getArchiveNum(self, x):
        try:
            idx = int(x.split("_")[3])
            return idx
        except:
            return 0

    def CheckDirForFiles(self, fpath):
        fileList = [
            f for f in os.listdir(fpath) if os.path.isfile(os.path.join(fpath, f))
        ]
        if fileList:
            return True
        else:
            return False

    def MsgToPDF(self, docpath, fpath, fname):
        pdfFileFormatCode = 17
        try:
            doc = self.objWord.Documents.Open(docpath)
            doc.SaveAs(os.path.join(fpath, fname), FileFormat=pdfFileFormatCode)
        except Exception as e:
            raise RuntimeError(f"Failed to convert {fname} to PDF. Error: {e}")
        finally:
            doc.Close()
            os.remove(docpath)


# * Main
class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("Email Archive Tool")
        # self.iconbitmap(os.path.join(os.getcwd(), "assets/email.ico"))

        view = AppView(self)
        view.grid(row=0, column=0, padx=10, pady=10)

        controller = AppController(UserInput, view)
        view.setController(controller)


if __name__ == "__main__":
    app = App()
    app.mainloop()
