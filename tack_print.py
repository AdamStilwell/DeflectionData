def worksheet_raw_print(workbook, my_tack):
    worksheet = workbook.add_worksheet(my_tack.sample_name)
    cell_format_string = workbook.add_format({"bold": True})
    cell_format_string.set_align("right")
    worksheet.set_column("A:A", 10)
    headers = ["Sample name", "Test force", "Test speed", "Test step"]
    worksheet.write_column("A1", headers, cell_format_string)

    worksheet.write(0, 1, my_tack.sample_name, cell_format_string)
    worksheet.write(1, 1, str(my_tack.test_force) + "N")
    worksheet.write(2, 1, str(my_tack.test_speed) + "um")
    worksheet.write(3, 1, str(my_tack.time_step) + "s")

    j = 0
    for array in my_tack.full_data_array:
        worksheet.write(6, j, my_tack.headers[-2][j], cell_format_string)
        start = 7
        for i in range(len(array)):
            worksheet.write(start, j, array[i])
            start += 1
        j += 1


def print_summary_worksheet(worksheet, my_tack, number_of_sheets):
    worksheet.write(number_of_sheets, 0, my_tack.sample_name)
    worksheet.write(number_of_sheets, 1, my_tack.max_load)
    worksheet.write(number_of_sheets, 2, my_tack.width)
    worksheet.write(number_of_sheets, 4, my_tack.g1c)
    worksheet.write(number_of_sheets, 5, my_tack.detach_pressure)
