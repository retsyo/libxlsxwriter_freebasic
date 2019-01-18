/'
 * Example of using libxlsxwriter for writing large files in constant memory
 * mode.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

function main() as Integer

    Dim as lxw_row_t row
    Dim as lxw_col_t col
    Dim as lxw_row_t max_row = 1000
    Dim as lxw_col_t max_col = 50

    /' Set the worksheet options. '/
    Dim as lxw_workbook_options options
    options.constant_memory = LXW_TRUE
    options.tmpdir = NULL

    /' Create a new workbook with options. '/
    Dim as lxw_workbook  ptr workbook  => workbook_new_opt("constant_memory.xlsx", @options)
    Dim as lxw_worksheet ptr worksheet => workbook_add_worksheet(workbook, NULL)

    for row = 0 to max_row-1
        for col = 0 to max_col-1
            worksheet_write_number(worksheet, row, col, 123.45, NULL)
        next
    next

    return workbook_close(workbook)
End Function

main()
