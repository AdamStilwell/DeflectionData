def worksheet_raw_print(workbook, my_deflection):
    worksheet = workbook.add_worksheet(my_deflection.sample_name)
    cell_format_string = workbook.add_format({"bold": True})
    cell_format_string.set_align("right")
    worksheet.set_column("A:A", 10)
    headers = ["Sample name", "Density", "Test force", "Test speed", "Test step"]
    worksheet.write_column("A1", headers, cell_format_string)

    worksheet.write(0, 1, my_deflection.sample_name, cell_format_string)
    worksheet.write(1, 1, my_deflection.density)
    worksheet.write(2, 1, str(my_deflection.test_force) + "N")
    worksheet.write(3, 1, str(my_deflection.test_speed) + "um")
    worksheet.write(4, 1, str(my_deflection.time_step) + "s")

    j = 0
    for array in my_deflection.full_data_array:
        worksheet.write(6, j, my_deflection.headers[-2][j], cell_format_string)
        start = 7
        for i in range(len(array)):
            worksheet.write(start, j, array[i])
            start += 1
        j += 1

    # Pressure vs Gap
    pressure_gap_chart = new_chart_creation(workbook=workbook,
                                            data_length=len(my_deflection.deflection_array),
                                            sample_name=my_deflection.sample_name,
                                            x_col=4,
                                            y_col=1,
                                            x_name="Pressure (MPa)",
                                            y_name="Gap (mm)")
    insert_chart_into_worksheet(worksheet=worksheet,
                                chart=pressure_gap_chart,
                                insert_location="K1",
                                x_size=0.65,
                                y_size=0.75)
    # Gap vs Load
    gap_load_chart = new_chart_creation(workbook=workbook,
                                        data_length=len(my_deflection.deflection_array),
                                        sample_name=my_deflection.sample_name,
                                        x_col=1,
                                        y_col=2,
                                        x_name="Displacement (mm)",
                                        y_name="Load (N)")
    insert_chart_into_worksheet(worksheet=worksheet,
                                chart=gap_load_chart,
                                insert_location="P1",
                                x_size=0.65,
                                y_size=0.75)

    # Pressure vs Deflection
    pressure_deflection_chart = new_chart_creation(workbook=workbook,
                                                   data_length=len(my_deflection.deflection_array),
                                                   sample_name=my_deflection.sample_name,
                                                   x_col=4,
                                                   y_col=3,
                                                   x_name="Pressure (MPa)",
                                                   y_name="Deflection (%)")
    insert_chart_into_worksheet(worksheet=worksheet,
                                chart=pressure_deflection_chart,
                                insert_location="K12",
                                x_size=0.65,
                                y_size=0.75)

    # Deflection vs Pressure
    deflection_pressure_chart = new_chart_creation(workbook=workbook,
                                                   data_length=len(my_deflection.deflection_array),
                                                   sample_name=my_deflection.sample_name,
                                                   x_col=3,
                                                   y_col=4,
                                                   x_name="Strain (%)",
                                                   y_name="Stress (MPa)")
    insert_chart_into_worksheet(worksheet=worksheet,
                                chart=deflection_pressure_chart,
                                insert_location="P12",
                                x_size=0.65,
                                y_size=0.75)

    # 1/h vs Pressure
    h_pressure_chart = new_chart_creation(workbook=workbook,
                                          data_length=len(my_deflection.deflection_array),
                                          sample_name=my_deflection.sample_name,
                                          x_col=7,
                                          y_col=4,
                                          x_name="1/h*dh/dt (1/s)",
                                          y_name="Pressure (MPa)")
    h_pressure_chart.add_series({
        "categories": [my_deflection.sample_name, 7, 7, len(my_deflection.deflection_array), 7],
        "values": [my_deflection.sample_name, 7, 8, len(my_deflection.deflection_array), 8],
        "name": my_deflection.sample_name,
        "line": {"width": 0.75, "color": "orange"},
    })
    insert_chart_into_worksheet(worksheet=worksheet,
                                chart=h_pressure_chart,
                                insert_location="K23",
                                x_size=1.25,
                                y_size=1.25)
    power_law_strings = ["Amplitude", "Power Law Index", "Offset"]
    worksheet.write_column("D1", power_law_strings, cell_format_string)

    # write the power law variables
    for i in range(len(my_deflection.power_law_values[0])):
        worksheet.write(("E" + str(i+1)), str(my_deflection.power_law_values[0][i]))

    # write all model stuff
    worksheet.write("I7", "model", cell_format_string)
    create_model(worksheet=worksheet,
                 data_length=len(my_deflection.deflection_array))
    worksheet.autofit()


def create_model(worksheet, data_length):
    offset = 8
    for i in range(0, data_length):
        worksheet.write("I" + str(i+offset), "=$E$1*(H" + str(i + offset) +
                        "-$E$3)^$E$2")


def insert_chart_into_worksheet(worksheet, chart, insert_location, x_size, y_size):
    worksheet.insert_chart(insert_location, chart,
                           {"x_scale": x_size, "y_scale": y_size})


def new_chart_creation(workbook, data_length, sample_name, x_col, y_col, x_name, y_name):
    chart = workbook.add_chart({"type": "scatter", "subtype": "smooth"})
    insert_values_into_chart(chart, data_length, x_col, y_col, sample_name)
    chart.set_legend({"none": True})
    chart.set_x_axis({
        "name": x_name,
        "name_font": {"size": 8, "bold": True}
    })
    chart.set_y_axis({
        "name": y_name,
        "name_font": {"size": 8, "bold": True}
    })
    return chart


def insert_values_into_chart(chart, data_length, x_col, y_col, sample_name):
    # find out how to make graphs
    chart.add_series({
        "categories": [sample_name, 7, x_col, data_length, x_col],
        "values": [sample_name, 7, y_col, data_length, y_col],
        "name": sample_name,
        "line": {"width": 0.75}
    })


def print_summary_worksheet(worksheet, my_deflection, number_of_sheets):
    worksheet.write(number_of_sheets, 0, my_deflection.sample_name)
    worksheet.write(number_of_sheets, 1, my_deflection.max_load)
    worksheet.write(number_of_sheets, 2, my_deflection.width)
    worksheet.write(number_of_sheets, 3, my_deflection.density)
    worksheet.write(number_of_sheets, 4, my_deflection.g1c)
    worksheet.write(number_of_sheets, 5, my_deflection.detach_pressure)
    worksheet.write(number_of_sheets, 6, my_deflection.max_deflection)
    worksheet.write(number_of_sheets, 7, my_deflection.minimum_gap)
    worksheet.write(number_of_sheets, 8, my_deflection.strain_to_break)

    # power law headings
    # "Amplitude"
    worksheet.write(number_of_sheets, 9, my_deflection.power_law_values[0][0])
    # "Power Law Index"
    worksheet.write(number_of_sheets, 10, my_deflection.power_law_values[0][1])
    # "Offset"
    worksheet.write(number_of_sheets, 11, my_deflection.power_law_values[0][2])
