/'
 * Example of writing some data to a simple Excel file using libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

function main() as Integer

    Dim As lxw_workbook  ptr workbook  = workbook_new("hello_world.xlsx")
    Dim As lxw_worksheet ptr worksheet = workbook_add_worksheet(workbook, NULL)

    worksheet_write_string(worksheet, 0, 0, "Hello", NULL)
    worksheet_write_number(worksheet, 1, 0, 123, NULL)

    workbook_close(workbook)

    return 0
End Function

main()
