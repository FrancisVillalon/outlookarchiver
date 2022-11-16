import os
import string

import pandas as pd
import win32com.client

from AppLogger import logger


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
