from tkinter import *
from tkinter import filedialog, ttk
from convert import changeFile, saveFile
import time
import os
import threading

#################################################################################
####           #This program converts a PPTX file to DOCX file#              ####
####                                                                         ####
####  This script sets up a gui using Python's Tkinter module and calls the  #### 
####  "convert.py" file to handle all logic associated with this program.    ####
#################################################################################


root = Tk()
root.geometry('600x400')
root.title("PowerPoint to Docx Conversion")

######### Progess Bar logic ###########
def bar():
    
    progress =  root.nametowidget("progressBar")

    root.nametowidget("progressLabel").place(relx=0.5, rely=0.4, anchor=CENTER)
    progress.place(relx=0.5, rely=0.5, anchor=CENTER)

    progress['value']=12.5
    root.update_idletasks()
    time.sleep(0.3)
    progress['value']=25
    root.update_idletasks()
    time.sleep(0.3)
    progress['value']=37.5
    root.update_idletasks()
    time.sleep(0.3)
    progress['value']=50
    root.update_idletasks()
    time.sleep(0.3)
    progress['value']=62.5
    root.update_idletasks()
    time.sleep(0.3)
    progress['value']=75
    root.update_idletasks()
    time.sleep(0.3)
    progress['value']=87.5
    root.update_idletasks()
    time.sleep(0.3)
    progress['value']=100

    progress.destroy()
    root.nametowidget("progressLabel").destroy()

    root.nametowidget("completeLabel").place(relx=0.5, rely=0.3, anchor=CENTER)
    root.nametowidget("browse_files").place(relx=0.5, rely=0.4, anchor=CENTER)
    root.nametowidget("folderEntry").place(relx=0.5, rely=0.5, anchor=CENTER)


########## Browse files to open and handles the widgets associated with it ########
def openfile():
    global file_location

    root.filename = filedialog.askopenfilename(initialdir="c:/Users/dougm", title="Selct a file", filetypes=(("pptx files","*.pptx"),("all files","*.*")))

    file_location = root.filename

    name = f"Convert {os.path.basename(file_location)}"

    root.nametowidget("fileEntry").delete(0, END)
    root.nametowidget("fileEntry").insert(0,file_location)
    
    if root.nametowidget("convertButton").winfo_ismapped() == False:
        root.nametowidget("convertButton").place(relx=0.5, rely=0.6, anchor=CENTER)
        root.nametowidget("convertButton")["text"]= name

    else:
        root.nametowidget("convertButton")["text"]= name


########### Browse folders to store new Microsoft Word file and handles the widgets associated with it ###########
def getFileDirectory():
    global fileDirectory

    root.directory = filedialog.askdirectory()

    fileDirectory = root.directory

    root.nametowidget("folderEntry").delete(0,END)
    root.nametowidget("folderEntry").insert(0,fileDirectory)

    if root.nametowidget("save_file").winfo_ismapped() == False:
        root.nametowidget("save_file").place(relx=0.5, rely=0.7, anchor=CENTER)

    if root.nametowidget("warningLabel").winfo_ismapped() == False:
        root.nametowidget("warningLabel").place(relx=0.5, rely=0.6, anchor=CENTER)


########### Calls the function to convert PowerPoint file to Word file and removes all widgets to display progress bar ############
def convertfile():
    
    global docxFile

    docxFile = changeFile(file_location)

    root.nametowidget("convertButton").destroy() 
    root.nametowidget("introLabel").destroy() 
    root.nametowidget("select_value").destroy() 
    root.nametowidget("fileEntry").destroy()

    threading.Thread(target=bar).start()


########### Calls the function to save the new Word file in selected folder and handles widgets ############
def save():
    saveFile(docxFile, fileDirectory)

    root.nametowidget("completeLabel").destroy()
    root.nametowidget("browse_files").destroy()
    root.nametowidget("save_file").destroy()
    root.nametowidget("folderEntry").destroy()
    root.nametowidget("warningLabel").destroy()

    root.nametowidget("savedLabel").place(relx=0.5, rely=0.4, anchor=CENTER)
    root.nametowidget("closeWindow").place(relx=0.5, rely=0.5, anchor=CENTER)
 



#############  sets widgets ###################

introLabel = Label(root, text="Convert your PowerPoint file to a Microsoft Word file.", name="introLabel")

warningLabel = Label(root, text="(Warning: If you have a Docx file named 'Conversion.docx' in selected folder it will be overwritten.)", name="warningLabel")

completeLabel = Label(root, text="Conversion successful! select folder to store new file.", name="completeLabel")

progressLabel = Label(root, text="Converting...", name="progressLabel")

fileEntry = Entry(root, name="fileEntry", width=50)

convertButton = Button(root, command=convertfile, name='convertButton')

folderEntry = Entry(root, name="folderEntry", width=50)

select_file = Button(root, text="Open file", command=openfile, name="select_value")

browse_files = Button(root, text="browse files", command=getFileDirectory, name="browse_files")

save_file = Button(root, text="Save", command=save, name="save_file")

savedLabel = Label(root, text="Save complete!", name="savedLabel")

closeWindow = Button(root, text="Close", command=root.destroy, name="closeWindow")

progressBar = ttk.Progressbar(root,orient=HORIZONTAL,length=350,mode='determinate', name="progressBar")

############# displays widgets at program start ##################

introLabel.place(relx=0.5, rely=0.3, anchor=CENTER)

select_file.place(relx=0.5, rely=0.4, anchor=CENTER)

fileEntry.place(relx=0.5, rely=0.5, anchor=CENTER)




############# Start program ####################
root.mainloop()