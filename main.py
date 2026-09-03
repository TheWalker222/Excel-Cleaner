import os
from tkinter import filedialog
import tkinter as tk
import customtkinter as ctk
from CleanExcel import clean_excel_file
import threading
files = {}
window = ctk.CTk()
window.geometry("800x500")
window.title("Clean Excel")
buttonFont = ctk.CTkFont(family="Arial", size=18, weight="bold")
labelFont = ctk.CTkFont(family="Arial", size=30, weight="bold")
enterFrame = ctk.CTkFrame(window, height=50, width=300)
enterFrame.grid(padx=5, pady=5, row=0, column=0)
enterLabel = ctk.CTkLabel(enterFrame, text="CLEAN YOUR EXCEL FILES", font=labelFont)
enterLabel.pack(padx=5, pady=5)
resultFrame = ctk.CTkFrame(window, height=400, width=200)
resultFrame.grid(padx=5, pady=5, row=1, column=1)
deleteEmptyVariable = tk.BooleanVar(value=True)
deleteEmptyCheckBox = ctk.CTkCheckBox(resultFrame, text="Delete Empty Rows", variable=deleteEmptyVariable, font=buttonFont)
deleteDuplicatesVariable = tk.BooleanVar(value=True)
deleteDuplicatesCheckBox = ctk.CTkCheckBox(resultFrame, text="Delete Duplicate Rows", variable=deleteDuplicatesVariable, font=buttonFont)
deleteEmptyCheckBox.pack(padx=5, pady=5, anchor="w")
deleteDuplicatesCheckBox.pack(padx=5, pady=5, anchor="w")
fileFrame = ctk.CTkFrame(window, height=300, width=200)
fileFrame.grid(padx=5, pady=5, row=1, column=0)
buttonFrame = ctk.CTkFrame(window, height=70, width=250)
buttonFrame.grid(padx=5, pady=5, column=1,row=0)
fileList = tk.Listbox(fileFrame, height=10, width=30,font=("Arial",24), selectmode='multiple')
fileList.pack()
addButton = ctk.CTkButton(buttonFrame,font=buttonFont, text="ADD FILES",height=60, width=100, fg_color="#878383", hover_color="#454343", command=lambda: add())
addButton.grid(row=0, column=0, padx=5, pady=5)
doButton = ctk.CTkButton(buttonFrame, text="DO",font=buttonFont, height=60, width=100, fg_color="#878383", hover_color="#454343", command=lambda: finishStart())
doButton.grid(row=0, column=1, padx=5, pady=5)
def add():
    global files
    selected = filedialog.askopenfilenames()
    for file in selected:
        if file not in files and file.lower().endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
            files[file] = os.path.basename(file)
    update()
def update():
    fileList.delete(0,tk.END)
    for file in files.values():
        fileList.insert(tk.END, file)
def finishStart():
    thread = threading.Thread(target=finish, daemon=True)
    thread.start()
    addButton.configure(state="disabled")
    doButton.configure(state="disabled")
    fileList.delete(0 ,tk.END)
    fileList.insert(tk.END, "Processing...")
def done():
    global files
    addButton.configure(state="normal")
    doButton.configure(state="normal")
    fileList.delete(0 ,tk.END)
    fileList.insert(tk.END, "Done!")
    files = {}
def finish():
    global files
    try:
        if not files:
            return
        clean_excel_file(list(files.keys()), delete_duplicates=deleteDuplicatesVariable.get(), delete_empty=deleteEmptyVariable.get())
    finally:
        window.after(0, done)
def delete(event):
    selected = fileList.curselection()
    if selected:
        for i in reversed(selected):
            files.pop(list(files.keys())[i])
    update()
fileList.bind("<Double-Button-1>", delete)
fileList.bind("<Delete>", delete)
fileList.bind("<BackSpace>", delete)
window.mainloop()