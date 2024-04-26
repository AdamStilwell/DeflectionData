import customtkinter as ctk
import xlsxwriter
import os
from tkinter import filedialog

from DeflectionData.deflection import Deflection
import ExcelPrint

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

root = ctk.CTk()
root.geometry("500x350")


def upload():
    global file_path
    global save_file
    if entry1.get() == "":
        save_file = "results.xlsx"
    else:
        save_file = entry1.get() + ".xlsx"
    file_path = filedialog.askopenfilenames(filetypes=[("csv file", ".csv")])
    # os.listdir(return_file_path)


def save_location():
    global save_file_location
    save_file_location = filedialog.askdirectory()


frame = ctk.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

label = ctk.CTkLabel(master=frame, text="Select folder of sample data", font=("Helvetica", 24))
label.pack(pady=12, padx=10)

entry1 = ctk.CTkEntry(master=frame, placeholder_text="Save File Name")
entry1.pack(pady=12, padx=10)

button = ctk.CTkButton(master=frame, text="Select Save Location", command=save_location)
button.pack(pady=12, padx=10)

button2 = ctk.CTkButton(master=frame, text="Select Samples", command=upload)
button2.pack(pady=12, padx=10)

if __name__ == "__main__":
    root.mainloop()
    workbook = xlsxwriter.Workbook(save_file_location + "/" + save_file)

    for filename in file_path:
        my_deflection = Deflection(filename=filename, headers_num=8)
        ExcelPrint.worksheet_raw_print(workbook, my_deflection)

    # print out data to Excel sheet here

    workbook.close()

