/'
 * A simple program to write some data to an Excel file using the
 * libxlsxwriter library.
 *
 * This program is shown, with explanations, in Tutorial 1 of the
 * libxlsxwriter documentation.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

' ~ #include "xlsxwriter.bi"
#include "auto_xlsxwriter.bi"

' ~ /* Some data we want to write to the worksheet. */
type expense
    item As String * 32
    cost As Integer
end Type

Dim Shared As expense expenses(0 to 3) => { ("Rent", 1000), ("Gas",   100), ("Food",  300), ("Gym",    50) }


function main() As Integer

    'Rem Create a workbook and add a worksheet. */
    dim as lxw_workbook  ptr workbook  => workbook_new("tutorial01.xlsx")
    dim as lxw_worksheet ptr worksheet => workbook_add_worksheet(workbook, NULL)

    'Rem Start from the first cell. Rows and columns are zero indexed. */
    dim as integer row = 0
    dim as integer col = 0

    'Rem Iterate over the data and write it out element by element. */
    for row = 0 to 3
        worksheet_write_string(worksheet, row, col,     expenses(row).item, NULL)
        worksheet_write_number(worksheet, row, col + 1, expenses(row).cost, NULL)
    next

    'Rem Write a total using a formula. */
    worksheet_write_string (worksheet, row, col,     "Total",       NULL)
    worksheet_write_formula(worksheet, row, col + 1, "=SUM(B1:B4)", NULL)

    'Rem Save the workbook and free any allocated memory. */
    return workbook_close(workbook)

end function

main()
