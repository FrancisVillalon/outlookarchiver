import os


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
