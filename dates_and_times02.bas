/'
 * Example of writing dates and times in Excel using an lxw_datetime struct
 * and date formatting.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

function main() as Integer

    /' A datetime to display. '/
    dim as lxw_datetime datetime = (2013, 2, 28, 12, 0, 0.0)


    /' Create a new workbook and add a worksheet. '/
    var workbook  = workbook_new("dates_and_times02.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)

    /' Add a format with date formatting. '/
    var format    = workbook_add_format(workbook)
    format_set_num_format(format, "mmm d yyyy hh:mm AM/PM")

    /' Widen the first column to make the text clearer. '/
    worksheet_set_column(worksheet, 0, 0, 20, NULL)

    /' Write the datetime without formatting. '/
    worksheet_write_datetime(worksheet, 0, 0, @datetime, NULL  )  Rem 41333.5

    /' Write the datetime with formatting. '/
    worksheet_write_datetime(worksheet, 1, 0, @datetime, format)  Rem Feb 28 2013 12:00 PM

    return workbook_close(workbook)
End Function

main()