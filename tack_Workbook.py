import xlsxwriter


class Workbook:

    def create_worksheet(self, title):
        worksheet = self.workbook.add_worksheet(title)
        return worksheet

    def create_chart(self, worksheet, location, x_name, y_name, chart_type, chart_subtype, chart_scale):
        chart = self.workbook.add_chart({"type": chart_type, "subtype": chart_subtype})
        worksheet.insert_chart(location, chart,
                               {"x_scale": chart_scale, "y_scale": chart_scale})
        chart.set_legend({"position": "top"})
        chart.set_x_axis({
            "name": x_name,
            "name_font": {"size": 8, "bold": True}
        })
        chart.set_y_axis({
            "name": y_name,
            "name_font": {"size": 8, "bold": True}
        })
        return chart

    def __init__(self, save_file_location, save_file, number_of_samples):
        self.workbook = xlsxwriter.Workbook(save_file_location + "/" + save_file)
        # create summary worksheet with all its headings
        self.worksheet_summary = self.workbook.add_worksheet("Summary")
        self.worksheet_summary.freeze_panes(1,0)
        self.worksheet_summary.set_column("A:A", 15)
        headers_array = ["Sample name", "Peak Load (N)", "Thickness (mm)",
                         "G1c (J/m2)", "Peak Detach Pressure (MPa)"]
        cell_format_string = self.workbook.add_format({"bold": True})
        cell_format_string.set_align("right")
        self.worksheet_summary.write_row("A1", headers_array, cell_format_string)

        self.worksheet_summary.autofit()

        self.worksheet_force_displacement = self.create_worksheet("Force-Displacement")
        self.force_displacement_chart = self.create_chart(worksheet=self.worksheet_force_displacement,
                                                          location="A1",
                                                          x_name="Gap between platens (mm)",
                                                          y_name="Force(N)",
                                                          chart_type="scatter",
                                                          chart_subtype="smooth",
                                                          chart_scale=1.5)

        self.worksheet_pressure_compression_rate = self.create_worksheet("Pressure-Compression Rate")
        self.pressure_compression_rate_chart = self.create_chart(worksheet=self.worksheet_pressure_compression_rate,
                                                                 location="A1",
                                                                 x_name="1/h*dh/dt (1/s)",
                                                                 y_name="Pressure (MPa)",
                                                                 chart_type="scatter",
                                                                 chart_subtype="smooth",
                                                                 chart_scale=1.5)