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
                                insert_location="J1")
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
                                insert_location="O1")

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
                                insert_location="J12")

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
                                insert_location="O12")

    # 1/h vs Pressure
    worksheet.autofit()


def insert_chart_into_worksheet(worksheet, chart, insert_location):
    worksheet.insert_chart(insert_location, chart,
                           {"x_scale": 0.65, "y_scale": 0.75})


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
    worksheet.write(number_of_sheets, 9, "")
    # "Power Law Index"
    worksheet.write(number_of_sheets, 10, "")
    # "Offset"
    worksheet.write(number_of_sheets, 11, "")
