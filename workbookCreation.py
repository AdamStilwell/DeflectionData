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

    def summary_sheet_charts(self, worksheet, number_of_samples, chart_offset, chart_loc, x_name, y_name, title,
                             sheet_name, x_col, y_col):
        chart = self.create_chart(worksheet=worksheet,
                                  location=(chart_loc + str(number_of_samples + chart_offset)),
                                  x_name=x_name,
                                  y_name=y_name,
                                  chart_type="scatter",
                                  chart_subtype="straight_with_markers",
                                  chart_scale=1)
        chart.add_series({
            "categories": [sheet_name, 1, x_col, number_of_samples, x_col],
            "values": [sheet_name, 1, y_col, number_of_samples, y_col],
            "name": title,
            "line": {"none": True}
        })

    def __init__(self, save_file_location, save_file, number_of_samples):
        self.workbook = xlsxwriter.Workbook(save_file_location + "/" + save_file)
        # create summary worksheet with all its headings
        self.worksheet_summary = self.workbook.add_worksheet("Summary")
        self.worksheet_summary.freeze_panes(1,0)
        self.worksheet_summary.set_column("A:A", 15)
        headers_array = ["Sample name", "Peak Load (N)", "Thickness (mm)", "Density (g/cc)",
                         "G1c (J/m2)", "Peak Detach Pressure (MPa)", "Maximum % Deflection",
                         "Minimum Gap (mm)", "Distance to Break (mm)", "Amplitude",
                         "Power Law Index", "Offset", "FTA-4 equiv.", "G'20 C", "G' 200 C", "Experiment Variable 1",
                         "Experiment Variable 2"]
        cell_format_string = self.workbook.add_format({"bold": True})
        cell_format_string.set_align("right")
        self.worksheet_summary.write_row("A1", headers_array, cell_format_string)

        # print all the summary charts
        # G1c vs Thickness
        self.summary_sheet_charts(worksheet=self.worksheet_summary,
                                  number_of_samples=number_of_samples,
                                  chart_offset=3,
                                  chart_loc="C",
                                  x_name="Thickness (mm)",
                                  y_name="G1c(J/m2)",
                                  title="G1c vs Thickness",
                                  sheet_name="Summary",
                                  x_col=2,
                                  y_col=4)
        # Peak Detach Pressure vs Thickness
        self.summary_sheet_charts(worksheet=self.worksheet_summary,
                                  number_of_samples=number_of_samples,
                                  chart_offset=3,
                                  chart_loc="H",
                                  x_name="Thickness (mm)",
                                  y_name="Peak Detach Pressure (MPa)",
                                  title="Peak Detach Pressure vs Thickness",
                                  sheet_name="Summary",
                                  x_col=2,
                                  y_col=5)
        # Minimum Gap vs Thickness
        self.summary_sheet_charts(worksheet=self.worksheet_summary,
                                  number_of_samples=number_of_samples,
                                  chart_offset=18,
                                  chart_loc="C",
                                  x_name="Thickness (mm)",
                                  y_name="Minimum Gap (mm)",
                                  title="Minimum Gap vs Thickness",
                                  sheet_name="Summary",
                                  x_col=2,
                                  y_col=7)
        # Maximum % Deflection vs Thickness
        self.summary_sheet_charts(worksheet=self.worksheet_summary,
                                  number_of_samples=number_of_samples,
                                  chart_offset=18,
                                  chart_loc="H",
                                  x_name="Thickness (mm)",
                                  y_name="Maximum % Deflection",
                                  title="Maximum % Deflection vs Thickness",
                                  sheet_name="Summary",
                                  x_col=2,
                                  y_col=6)

        # Experimental Variable 1 vs Thickness
        self.summary_sheet_charts(worksheet=self.worksheet_summary,
                                  number_of_samples=number_of_samples,
                                  chart_offset=3,
                                  chart_loc="M",
                                  x_name="Experimental Variable 1",
                                  y_name="Thickness (mm)",
                                  title="Thickness vs Experimental Variable 1",
                                  sheet_name="Summary",
                                  x_col=15,
                                  y_col=2)

        # Experimental Variable 1 vs G1c
        self.summary_sheet_charts(worksheet=self.worksheet_summary,
                                  number_of_samples=number_of_samples,
                                  chart_offset=3,
                                  chart_loc="S",
                                  x_name="Experimental Variable 1",
                                  y_name="G1c",
                                  title="G1c vs Experimental Variable 1",
                                  sheet_name="Summary",
                                  x_col=15,
                                  y_col=4)

        # Experimental Variable 1 vs Peak Detach
        self.summary_sheet_charts(worksheet=self.worksheet_summary,
                                  number_of_samples=number_of_samples,
                                  chart_offset=18,
                                  chart_loc="M",
                                  x_name="Experimental Variable 1",
                                  y_name="Peak Detach (MPa)",
                                  title="Peak Detach vs Experimental Variable 1",
                                  sheet_name="Summary",
                                  x_col=15,
                                  y_col=5)

        # Experimental Variable 1 vs Max Deflection
        self.summary_sheet_charts(worksheet=self.worksheet_summary,
                                  number_of_samples=number_of_samples,
                                  chart_offset=18,
                                  chart_loc="S",
                                  x_name="Experimental Variable 1",
                                  y_name="Max Deflection (%)",
                                  title="Thickness vs Experimental Variable 1",
                                  sheet_name="Summary",
                                  x_col=15,
                                  y_col=6)

        # Experimental Variable 1 vs Minimum Gap
        self.summary_sheet_charts(worksheet=self.worksheet_summary,
                                  number_of_samples=number_of_samples,
                                  chart_offset=33,
                                  chart_loc="M",
                                  x_name="Experimental Variable 1",
                                  y_name="Minimum Gap (mm)",
                                  title="Minimum Gap vs Experimental Variable 1",
                                  sheet_name="Summary",
                                  x_col=15,
                                  y_col=7)

        # Experimental Variable 1 vs Minimum Gap
        self.summary_sheet_charts(worksheet=self.worksheet_summary,
                                  number_of_samples=number_of_samples,
                                  chart_offset=33,
                                  chart_loc="S",
                                  x_name="Experimental Variable 1",
                                  y_name="Distance to Break (um)",
                                  title="Distance to Break vs Experimental Variable 1",
                                  sheet_name="Summary",
                                  x_col=15,
                                  y_col=8)

        self.worksheet_summary.autofit()

        self.worksheet_pressure_deflection = self.create_worksheet("Pressure Deflection 450N")
        self.pressure_deflection_chart = self.create_chart(worksheet=self.worksheet_pressure_deflection,
                                                           location="A1",
                                                           x_name="Pressure (psi)",
                                                           y_name="% Deflection",
                                                           chart_type="scatter",
                                                           chart_subtype="smooth",
                                                           chart_scale=1.5)

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
