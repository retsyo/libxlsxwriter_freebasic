/'
 * An example of how to hide rows and columns using the libxlsxwriter
 * library.
 *
 * In order to hide rows without setting each one, (of approximately 1 million
 * rows), Excel uses an optimization to hide all rows that don't have data.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

function main() as Integer

    /' Create a new workbook and add a worksheet. '/
    var workbook  = workbook_new("hide_row_col.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)
    Dim As lxw_row_t row

    /' Write some data. '/
    worksheet_write_string(worksheet, 0, 3, "Some hidden columns.", NULL)
    worksheet_write_string(worksheet, 7, 0, "Some hidden rows.",    NULL)

    /' Hide all rows without data. '/
    worksheet_set_default_row(worksheet, 15, LXW_TRUE)

    /' Set the height of empty rows that we do want to display even if it is '/
    /' the default height. '/
    for row = 1 to 5
        worksheet_set_row(worksheet, row, 15, NULL)
    next

    /' Columns can be hidden explicitly. This doesn't increase the file size. '/
    Dim As lxw_row_col_options options
    options.hidden = 1

    worksheet_set_column_opt(worksheet, COLS("G:XFD"), 8.43, NULL, @options)

    workbook_close(workbook)

    return 0
End Function

main()