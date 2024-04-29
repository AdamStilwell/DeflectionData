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


def worksheet_print_graphs(workbook):
    # find out how to make graphs
    print("Hello")


def print_summary_worksheet(worksheet):
    print("World")
