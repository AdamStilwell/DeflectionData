import xlsxwriter


class Workbook:
    def __init__(self, save_file_location, save_file):
        self.workbook = xlsxwriter.Workbook(save_file_location + "/" + save_file)
        # create summary worksheet with all its headings
        self.worksheet_summary = self.workbook.add_worksheet("Summary")
        self.worksheet_summary.set_column("A:A", 15)
        headers_array = ["Sample name", "Peak Load (N)", "Thickness (mm)", "Density (g/cc)",
                         "G1c (J/m2)", "Peak Detach Pressure (MPa)", "Maximum % Deflection",
                         "Minimum Gap (mm)", "Distance to Break (um)", "Amplitude",
                         "Power Law Index", "Offset", "FTA-4 equiv.", "G'20 C", "G' 200 C"]
        self.worksheet_summary.write_row("A1", headers_array)
        self.worksheet_summary.autofit()

        self.worksheet_pressure_deflection = self.workbook.add_worksheet("Pressure Deflection 450N")
        self.pressure_deflection_chart = self.workbook.add_chart({"type": "scatter"})
        self.worksheet_pressure_deflection.insert_chart("A1", self.pressure_deflection_chart,
                                                        {"x_scale": 1.5, "y_scale": 1.5})

        self.worksheet_force_displacement = self.workbook.add_worksheet("Force-Displacement")
        self.force_displacement_chart = self.workbook.add_chart({"type": "scatter"})
        self.worksheet_force_displacement.insert_chart("A1", self.force_displacement_chart,
                                                       {"x_scale": 1.5, "y_scale": 1.5})

        self.worksheet_pressure_compression_rate = self.workbook.add_worksheet("Pressure-Compression Rate")
        self.pressure_compression_rate_chart = self.workbook.add_chart({"type": "scatter"})
        self.worksheet_pressure_compression_rate.insert_chart("A1", self.pressure_compression_rate_chart,
                                                              {"x_scale": 1.5, "y_scale": 1.5})
