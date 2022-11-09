from tkinter import *
from tkinter import ttk
from tkinter import filedialog,messagebox,Text
import win32com.client
import os 
import string

obj = win32com.client.Dispatch("Outlook.Application")
outlook = obj.GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

# ui
gui = Tk()
gui.geometry("550x260")
gui.title("Email Archive Tool v0")

def get_archive_num(x):
    return int(x.split("_")[3])

def getFolderPath():
    folder_selected = filedialog.askdirectory()
    folderPath.set(folder_selected)

def export_emails():
    try:
        # ? vars
        sm_opt = sm_bool.get()
        fname = folderPath.get()
        selected = obj.ActiveExplorer().Selection
        bp = bp_num.get()
        ordernum = order_num.get()
        emailcat = email_cat.get()
        num_emails = len(list(selected))
        msg = f"Do you want to archive {num_emails} selected emails?"   

        # ! Error messages

        if not os.path.exists(fname):
            messagebox.showerror("Error","Export path does not exist.")
            return
        if not fname:
            messagebox.showerror("Error","Please select an export folder")
            return
        if not len(list(selected)) > 0:
            messagebox.showerror("Error","No emails selected.")
            return
        if not bp:
            messagebox.showerror("Error","Please set BP")
            return
        if not ordernum:
            messagebox.showerror("Error","Please set Order number")
            return
        if not emailcat:
            messagebox.showerror("Error","Please set Email category")
            return

        # ! sm opt 
        if sm_opt:
            ans=True
        else:
            ans = messagebox.askyesno("Confirm",msg)
        
        # ! export code
        if ans:
            # create custom folder 
            if len(custom_fname.get().translate(str.maketrans('','',string.punctuation)))>0:

                if not os.path.exists(os.path.join(fname,custom_fname.get())) and custom_fname:
                    os.mkdir(os.path.join(fname,custom_fname.get()))
                if custom_fname:
                    fname = os.path.join(fname,custom_fname.get())
            d = 100/len(list(selected))

            if len(os.listdir(fname))>0:
                offsetidx = max(list(map(get_archive_num,os.listdir(fname))))
            else:
                offsetidx = 1

            for idx,email in enumerate(selected):
                try:
                    email_path_msg = os.path.join(
                        fname,
                        (
                            f"{bp}_{ordernum}_{emailcat}_{idx+offsetidx}_"+
                            email.subject.translate(str.maketrans('','',string.punctuation))
                            +".msg"
                        )
                    )
                    # archive as msg
                    email.SaveAs(email_path_msg)
                    # save attachments
                    for idx,atch in enumerate(email.Attachments):
                        atch.SaveASFile(os.path.join(
                            fname,f"{bp}_{ordernum}_{emailcat}_{idx+offsetidx}__{atch.FileName}"
                            ))
                    pb["value"] = d*(idx+1)
                    gui.update_idletasks()
                except Exception as e:
                    print(f"Failed to save email {idx} due to Exception : {e}")
                    continue
            if not sm_opt:
                messagebox.showinfo(title="Success",message=f"{num_emails} emails have been successfully archived at path: \n\n {fname}")
        else:
            messagebox.showinfo(title="Cancelled",message="Operation Cancelled")
    except Exception as e:
        print(e)
        messagebox.showerror("Unexpected Error","Unexpected error, Please check if outlook is open and check the export path.")
        
# export path label
folderPath = StringVar()
sm_bool = BooleanVar()
email_cat = StringVar()
email_cat.set("cat1")

a = Label(gui ,text="Export Folder: ")
a.grid(row=0,column = 0,sticky='w')

# export path entry
E = Entry(gui,textvariable=folderPath)
E.grid(row=0,column=1,ipadx=100)

# export path browser
exportpathbtn = ttk.Button(gui, text="Browse Folder",command=getFolderPath)
exportpathbtn.grid(row=0,column=2,padx=5)


# progress bar
pb = ttk.Progressbar(gui, orient = HORIZONTAL, mode = 'determinate',length=100)
pb.grid(row=3,column=1,columnspan=3,sticky="ew",padx=(0,5))
Label(gui,text="Task Progress: ").grid(row=3,column=0,sticky='w')

# supress messages
Label(gui,text="Options: ").grid(row = 4,column=0,sticky="W")
Checkbutton(gui,text="Supress messages",variable=sm_bool).grid(row = 4,sticky="W",column=1)

# folder name
custom_fname = Entry(gui)
custom_fname.grid(row=5,column=1,pady=5,sticky='w')
Label(gui,text="Create New Folder :  ").grid(row=5,column=0,sticky='w')

# bp_num
bp_num = Entry(gui)
bp_num.grid(row=6,column=1,pady=5,sticky='w')
Label(gui,text="Enter BP Num : ").grid(row=6,column=0,sticky='w')

# order_num
order_num = Entry(gui)
order_num.grid(row=7,column=1,pady=5,sticky='w')
Label(gui,text="Enter Order Num : ").grid(row=7,column=0,sticky='w')

# email_cat

email_cat_menu = OptionMenu(gui,email_cat,"cat1","cat2","cat3")
Label(gui,text="Enter Email Category : ").grid(row=8,column=0,sticky='w')
email_cat_menu.grid(row=8,column=1,pady=5,sticky='w')

# execute btn
c = ttk.Button(gui ,text="Export", command=export_emails)
c.grid(row=9,column=0,pady=5,sticky='w')


gui.mainloop()