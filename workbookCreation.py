import xlsxwriter
import ExcelPrint


class Workbook:

    def create_worksheet(self, title):
        worksheet = self.workbook.add_worksheet(title)
        return worksheet

    def create_chart(self, worksheet, location):
        chart = self.workbook.add_chart({"type": "scatter", "subtype": "smooth"})
        worksheet.insert_chart(location, chart,
                               {"x_scale": 1.5, "y_scale": 1.5})
        return chart

    def __init__(self, save_file_location, save_file, number_of_samples):
        self.workbook = xlsxwriter.Workbook(save_file_location + "/" + save_file)
        # create summary worksheet with all its headings
        self.worksheet_summary = self.workbook.add_worksheet("Summary")
        self.worksheet_summary.set_column("A:A", 15)
        headers_array = ["Sample name", "Peak Load (N)", "Thickness (mm)", "Density (g/cc)",
                         "G1c (J/m2)", "Peak Detach Pressure (MPa)", "Maximum % Deflection",
                         "Minimum Gap (mm)", "Distance to Break (um)", "Amplitude",
                         "Power Law Index", "Offset", "FTA-4 equiv.", "G'20 C", "G' 200 C"]
        cell_format_string = self.workbook.add_format({"bold": True})
        cell_format_string.set_align("right")
        self.worksheet_summary.write_row("A1", headers_array, cell_format_string)
        self.worksheet_summary.autofit()

        self.worksheet_pressure_deflection = self.create_worksheet("Pressure Deflection 450N")
        self.pressure_deflection_chart = self.create_chart(worksheet=self.worksheet_pressure_deflection,
                                                           location="A1")

        self.worksheet_force_displacement = self.create_worksheet("Force-Displacement")
        self.force_displacement_chart = self.create_chart(worksheet=self.worksheet_force_displacement,
                                                          location="A1")

        self.worksheet_pressure_compression_rate = self.create_worksheet("Pressure-Compression Rate")
        self.pressure_compression_rate_chart = self.create_chart(worksheet=self.worksheet_pressure_compression_rate,
                                                                 location="A1")
