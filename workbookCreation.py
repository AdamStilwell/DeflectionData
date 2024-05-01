import xlsxwriter


class Workbook:
    def __init__(self, save_file_location, save_file):
        self.workbook = xlsxwriter.Workbook(save_file_location + "/" + save_file)
        # create summary worksheet with all its headings
        self.worksheet_summary = self.workbook.add_worksheet("Summary")
        self.worksheet_summary.set_column("A:A", 15)
        self.worksheet_summary.write(0, 0, "Sample name")
        self.worksheet_summary.write(0, 1, "Peak Load (N)")
        self.worksheet_summary.write(0, 2, "Thickness (mm)")
        self.worksheet_summary.write(0, 3, "Density (g/cc)")
        self.worksheet_summary.write(0, 4, "G1c (J/m2)")
        self.worksheet_summary.write(0, 5, "Peak Detach Pressure (MPa)")
        self.worksheet_summary.write(0, 6, "Maximum % Deflection")
        self.worksheet_summary.write(0, 7, "Minimum Gap (mm)")
        self.worksheet_summary.write(0, 8, "Distance to Break (um)")
        # power law headings
        self.worksheet_summary.write(0, 9, "Amplitude")
        self.worksheet_summary.write(0, 10, "Power Law Index")
        self.worksheet_summary.write(0, 11, "Offset")

        self.worksheet_summary.write(0, 12, "FTA-4 equiv.")
        self.worksheet_summary.write(0, 13, "G'20 C")
        self.worksheet_summary.write(0, 14, "G' 200 C")
        self.worksheet_summary.autofit()

        self.worksheet_pressure_deflection = self.workbook.add_worksheet("Pressure Deflection 450N")
        self.pressure_deflection_chart = self.workbook.add_chart({"type": "scatter"})

        self.worksheet_force_displacement = self.workbook.add_worksheet("Force-Displacement")
        self.worksheet_force_displacement.set_column("A:A", 10)

        self.worksheet_pressure_compression_rate = self.workbook.add_worksheet("Pressure-Compression Rate")
        self.worksheet_pressure_compression_rate.set_column("A:A", 10)