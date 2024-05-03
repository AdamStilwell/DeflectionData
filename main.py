import customtkinter as ctk
from tkinter import filedialog

from DeflectionData.deflection import Deflection
import ExcelPrint
import workbookCreation


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
    root.destroy()


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
    workbook_class = workbookCreation.Workbook(save_file_location=save_file_location,
                                               save_file=save_file,
                                               number_of_samples=len(file_path))
    number_of_sheets = 1
    for filename in file_path:
        my_deflection = Deflection(filename=filename,
                                   headers_num=8)
        ExcelPrint.worksheet_raw_print(workbook=workbook_class.workbook,
                                       my_deflection=my_deflection)
        ExcelPrint.print_summary_worksheet(worksheet=workbook_class.worksheet_summary,
                                           my_deflection=my_deflection,
                                           number_of_sheets=number_of_sheets)
        # Pressure-Deflection
        ExcelPrint.insert_values_into_chart(chart=workbook_class.pressure_deflection_chart,
                                            data_length=len(my_deflection.sample_load_array),
                                            x_col=5,
                                            y_col=3,
                                            sample_name=my_deflection.sample_name)
        # Force-Displacement
        ExcelPrint.insert_values_into_chart(chart=workbook_class.force_displacement_chart,
                                            data_length=len(my_deflection.sample_load_array),
                                            x_col=1,
                                            y_col=2,
                                            sample_name=my_deflection.sample_name)
        # Pressure-CompressionRate
        ExcelPrint.insert_values_into_chart(chart=workbook_class.pressure_compression_rate_chart,
                                            data_length=len(my_deflection.sample_load_array),
                                            x_col=7,
                                            y_col=4,
                                            sample_name=my_deflection.sample_name)
        number_of_sheets += 1

    # print out data to Excel sheet here

    workbook_class.workbook.close()

