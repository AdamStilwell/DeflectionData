import pandas as pd
import customtkinter as ctk
import xlsxwriter
from tkinter import filedialog

from DeflectionData.deflection import Deflection

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

root = ctk.CTk()
root.geometry("500x350")


def upload():
    return filedialog.askdirectory()


frame = ctk.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

label = ctk.CTkLabel(master=frame, text="Select folder of sample data", font=("Helvetica", 24))
label.pack(pady=12, padx=10)

# entry1 = ctk.CTkEntry(master=frame, placeholder_text="File to upload")
# entry1.pack(pady=12, padx=10)

button = ctk.CTkButton(master=frame, text="Select Folder", command=upload)
button.pack(pady=12, padx=10)


if __name__ == "__main__":
    # "CN7480_1.0_1.csv"
    # root.mainloop()

    filenames = ["CN7480_1.0_1.csv"]
    deflections = []

    workbook = xlsxwriter.Workbook("test.xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.set_column("A:A", 20)

    for filename in filenames:
        my_deflection = Deflection(filename=filename, headers_num=6)
        deflections.append(my_deflection)
        for x in range(len(my_deflection.time_array)):
            pos = str.join("A", str(x))
            worksheet.write(pos, my_deflection.headers[0])

    for deflection in deflections:
        deflection.compile_data()
    # file_name = file_name.split("/")[-1]
    # print out data to Excel sheet here

    workbook.close()

