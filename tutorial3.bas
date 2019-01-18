/'
 * A simple program to write some data to an Excel file using the
 * libxlsxwriter library.
 *
 * This program is shown, with explanations, in Tutorial 3 of the
 * libxlsxwriter documentation.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

/' Some data we want to write to the worksheet. '/
type expense
    item As String * 32
    cost As Integer
    datetime As lxw_datetime
end Type

Dim Shared As expense expenses(0 to 3) /'=> {  _
    ("Rent", 1000, { .year = 2013, .month = 1, .day = 13 } ),  _
    ("Gas",   100, { .year = 2013, .month = 1, .day = 14 } ),  _
    ("Food",  300, { .year = 2013, .month = 1, .day = 16 } ),  _
    ("Gym",    50, { .year = 2013, .month = 1, .day = 20 } ),  _
}'/


sub write_worksheet_data(byval worksheet as lxw_worksheet ptr, byval bold as lxw_format ptr)
    Dim As integer row, col
    Dim As uinteger data_(0 to 5, 0 to 2) = { _
        /' Three columns of data. '/ _
        {2, 10, 30}, _
        {3, 40, 60}, _
        {4, 50, 70}, _
        {5, 20, 50}, _
        {6, 10, 40}, _
        {7, 50, 30} _
    }

    worksheet_write_string(worksheet, CELL("A1"), "Number",  bold)
    worksheet_write_string(worksheet, CELL("B1"), "Batch 1", bold)
    worksheet_write_string(worksheet, CELL("C1"), "Batch 2", bold)

    for row = 0 to 5
        for col = 0 to 2
            worksheet_write_number(worksheet, row + 1, col, data_(row, col) , NULL)
        next
    next

end sub
function main() as Integer

    /' Create a workbook and add a worksheet. '/
    var workbook  = workbook_new("tutorial03.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)
    Dim As integer row, col, i


    /' Add a bold format to use to highlight cells. '/
    var bold = workbook_add_format(workbook)
    format_set_bold(bold)

    /' Add a number format for cells with money. '/
    var money = workbook_add_format(workbook)
    format_set_num_format(money, "$#,##0")

    /' Add an Excel date format. '/
    var date_format = workbook_add_format(workbook)
    format_set_num_format(date_format, "mmmm d yyyy")

    /' Adjust the column width. '/
    worksheet_set_column(worksheet, 0, 0, 15, NULL)

    /' Write some data header. '/
    worksheet_write_string(worksheet, row, col,     "Item", bold)
    worksheet_write_string(worksheet, row, col + 1, "Cost", bold)

    /' Iterate over the data and write it out element by element. '/
    for i = 0 to 3
        /' Write from the first cell below the headers. '/
        row = i + 1
        worksheet_write_string  (worksheet, row, col,      expenses(i).item,     NULL)
        worksheet_write_datetime(worksheet, row, col + 1, @expenses(i).datetime, date_format)
        worksheet_write_number  (worksheet, row, col + 2,  expenses(i).cost,     money)
    next

    /' Write a total using a formula. '/
    worksheet_write_string (worksheet, row + 1, col,     "Total",       bold)
    worksheet_write_formula(worksheet, row + 1, col + 2, "=SUM(C2:C5)", money)

    /' Save the workbook and free any allocated memory. '/
    return workbook_close(workbook)
End Function

main()