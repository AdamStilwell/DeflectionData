def worksheet_raw_print(workbook, my_deflection):
    worksheet = workbook.add_worksheet(my_deflection.sample_name)
    worksheet.set_column("A:A", 10)

    worksheet.write(0, 0, "Sample name")
    worksheet.write(0, 1, my_deflection.sample_name)
    worksheet.write(1, 0, "Density")
    worksheet.write(1, 1, my_deflection.density)
    worksheet.write(2, 0, "Test force")
    worksheet.write(2, 1, my_deflection.test_force)
    worksheet.write(3, 0, "Test speed")
    worksheet.write(3, 1, my_deflection.test_speed)
    worksheet.write(4, 0, "Test step")
    worksheet.write(4, 1, my_deflection.time_step)

    j = 0
    for array in my_deflection.full_data_array:
        worksheet.write(6, j, my_deflection.headers[-2][j])
        start = 7
        for i in range(len(array)):
            worksheet.write(start, j, array[i])
            start += 1
        j += 1


def worksheet_print_graphs(workbook, chart, x_axis, y_axis):
    # find out how to make graphs
    print("Hello")


def print_summary_worksheet(worksheet, my_deflection, number_of_sheets):
    worksheet.write(number_of_sheets, 0, my_deflection.sample_name)
    worksheet.write(number_of_sheets, 1, my_deflection.max_load)
    worksheet.write(number_of_sheets, 2, my_deflection.width)
    worksheet.write(number_of_sheets, 3, my_deflection.density)
    worksheet.write(number_of_sheets, 4, "G1c (J/m2)")
    worksheet.write(number_of_sheets, 5, my_deflection.detach_pressure)
    worksheet.write(number_of_sheets, 6, my_deflection.max_deflection)
    worksheet.write(number_of_sheets, 7, my_deflection.minimum_gap)
    worksheet.write(number_of_sheets, 8, "Distance to Break (um)")
    # power law headings
    worksheet.write(number_of_sheets, 9, "Amplitude")
    worksheet.write(number_of_sheets, 10, "Power Law Index")
    worksheet.write(number_of_sheets, 11, "Offset")

    worksheet.write(number_of_sheets, 12, "FTA-4 equiv.")
    worksheet.write(number_of_sheets, 13, "G'20 C")
    worksheet.write(number_of_sheets, 14, "G' 200 C")
