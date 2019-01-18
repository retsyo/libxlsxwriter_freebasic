/'
 * A simple program to write some data to an Excel file using the
 * libxlsxwriter library.
 *
 * This program is shown, with explanations, in Tutorial 2 of the
 * libxlsxwriter documentation.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

/' Some data we want to write to the worksheet. '/
type expense
    item as string *32
    cost as integer
end type

dim shared as expense expenses(...) = { _
    ("Rent", 1000),                    _
    ("Gas",   100),                     _
    ("Food",  300),                    _
    ("Gym",    50)                    _
}


function main() as Integer

    /' Create a workbook and add a worksheet. '/
    var workbook  = workbook_new("tutorial02.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)
    Dim As integer row = 0, col = 0, i

    /' Add a bold format to use to highlight cells. '/
    var bold = workbook_add_format(workbook)
    format_set_bold(bold)

    /' Add a number format for cells with money. '/
    var money = workbook_add_format(workbook)
    format_set_num_format(money, "$#,##0")

    /' Write some data header. '/
    worksheet_write_string(worksheet, row, col,     "Item", bold)
    worksheet_write_string(worksheet, row, col + 1, "Cost", bold)

    /' Iterate over the data and write it out element by element. '/
    for i = 0 to 3
        /' Write from the first cell below the headers. '/
        row = i + 1
        worksheet_write_string(worksheet, row, col,     expenses(i).item, NULL)
        worksheet_write_number(worksheet, row, col + 1, expenses(i).cost, money)
    next

    /' Write a total using a formula. '/
    worksheet_write_string (worksheet, row + 1, col,     "Total",       bold)
    worksheet_write_formula(worksheet, row + 1, col + 1, "=SUM(B2:B5)", money)

    /' Save the workbook and free any allocated memory. '/
    return workbook_close(workbook)
end function

main()