/'
 * Example of writing some data with numeric formatting to a simple Excel file
 * using libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

function main() as Integer

    /' Create a new workbook and add a worksheet. '/
    var workbook  = workbook_new("format_num_format.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)

    /' Widen the first column to make the text clearer. '/
    worksheet_set_column(worksheet, 0, 0, 30, NULL)

    /' Add some formats. '/
    var format01   = workbook_add_format(workbook)
    var format02   = workbook_add_format(workbook)
    var format03   = workbook_add_format(workbook)
    var format04   = workbook_add_format(workbook)
    var format05   = workbook_add_format(workbook)
    var format06   = workbook_add_format(workbook)
    var format07   = workbook_add_format(workbook)
    var format08   = workbook_add_format(workbook)
    var format09   = workbook_add_format(workbook)
    var format10   = workbook_add_format(workbook)
    var format11   = workbook_add_format(workbook)

    /' Set some example number formats. '/
    format_set_num_format(format01, "0.000")
    format_set_num_format(format02, "#,##0")
    format_set_num_format(format03, "#,##0.00")
    format_set_num_format(format04, "0.00")
    format_set_num_format(format05, "mm/dd/yy")
    format_set_num_format(format06, "mmm d yyyy")
    format_set_num_format(format07, "d mmmm yyyy")
    format_set_num_format(format08, "dd/mm/yyyy hh:mm AM/PM")
    format_set_num_format(format09, "0 " & chr(34) & "dollar and" & chr(34) & " .00 " & chr(34) & "cents" & chr(34) & "")

    /' Write data using the formats. '/
    worksheet_write_number(worksheet, 0, 0, 3.1415926, NULL) Rem      3.1415926
    worksheet_write_number(worksheet, 1, 0, 3.1415926, format01) Rem  3.142
    worksheet_write_number(worksheet, 2, 0, 1234.56,   format02) Rem  1,235
    worksheet_write_number(worksheet, 3, 0, 1234.56,   format03) Rem  1,234.56
    worksheet_write_number(worksheet, 4, 0, 49.99,     format04) Rem  49.99
    worksheet_write_number(worksheet, 5, 0, 36892.521, format05) Rem  01/01/01
    worksheet_write_number(worksheet, 6, 0, 36892.521, format06) Rem  Jan 1 2001
    worksheet_write_number(worksheet, 7, 0, 36892.521, format07) Rem  1 January 2001
    worksheet_write_number(worksheet, 8, 0, 36892.521, format08) Rem  01/01/2001 12:30 AM
    worksheet_write_number(worksheet, 9, 0, 1.87,      format09) Rem  1 dollar and .87 cents

    /' Show limited conditional number formats. '/
    format_set_num_format(format10, "[Green]General;[Red]-General;General")
    worksheet_write_number(worksheet, 10, 0, 123, format10) Rem  > 0 Green
    worksheet_write_number(worksheet, 11, 0, -45, format10) Rem  < 0 Red
    worksheet_write_number(worksheet, 12, 0,   0, format10) Rem  = 0 Default color

    /' Format a Zip code. '/
    format_set_num_format(format11, "00000")
    worksheet_write_number(worksheet, 13, 0, 1209, format11)

    /' Close the workbook, save the file and free any memory. '/
    return workbook_close(workbook)
End Function

main()