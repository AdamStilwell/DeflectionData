import xlsxwriter


class Workbook:
    def __init__(self, save_file_location, save_file):
        self.workbook = xlsxwriter.Workbook(save_file_location + "/" + save_file)
        # create summary worksheet with all its headings
        self.worksheet_summary = self.workbook.add_worksheet("Summary")
        self.worksheet_summary.set_column("A:A", 10)

        self.worksheet_pressure_deflection = self.workbook.add_worksheet("Pressure Deflection 450N")
        self.worksheet_pressure_deflection.set_column("A:A", 10)

        self.worksheet_force_displacement = self.workbook.add_worksheet("Force-Displacement")
        self.worksheet_force_displacement.set_column("A:A", 10)

        self.worksheet_pressure_compression_rate = self.workbook.add_worksheet("Pressure-Compression Rate")
        self.worksheet_pressure_compression_rate.set_column("A:A", 10)