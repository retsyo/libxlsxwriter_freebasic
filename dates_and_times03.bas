/'
 * Example of writing dates and times in Excel using different date formats.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

function main() as Integer

    /' A datetime to display. '/
    dim as lxw_datetime datetime = (2013, 1, 23, 12, 30, 5.123)

    var  row = 0
    var col = 0
    dim as integer          i

    /' Examples date and time formats. In the output file compare how changing
     * the format strings changes the appearance of the date.
     '/
    Dim as zstring * 32 date_formats(...) = { _
        "dd/mm/yy",                               _
        "mm/dd/yy",                               _
        "dd m yy",                                  _
        "d mm yy",                                 _
        "d mmm yy",                              _
        "d mmmm yy",                           _
        "d mmmm yyy",                         _
        "d mmmm yyyy",                       _
        "dd/mm/yy hh:mm",                  _
        "dd/mm/yy hh:mm:ss",              _
        "dd/mm/yy hh:mm:ss.000",       _
        "hh:mm",                                  _
        "hh:mm:ss",                             _
        "hh:mm:ss.000"                      _
    }

    /' Create a new workbook and add a worksheet. '/
    var workbook  = workbook_new("dates_and_times03.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)

    /' Add a bold format. '/
    dim As lxw_format  ptr   bold      = workbook_add_format(workbook)
    format_set_bold(bold)

    /' Write the column headers. '/
    worksheet_write_string(worksheet, row, col,     "Formatted date", bold)
    worksheet_write_string(worksheet, row, col + 1, "Format",         bold)

    /' Widen the first column to make the text clearer. '/
    worksheet_set_column(worksheet, 0, 1, 20, NULL)

    /' Write the same date and time using each of the above formats. '/
    for i = 0 to 13
        row += 1

        /' Create a format for the date or time.'/
        var format  = workbook_add_format(workbook)
        format_set_num_format(format, date_formats(i))
        format_set_align(format, LXW_ALIGN_LEFT)

        /' Write the datetime with each format. '/
        worksheet_write_datetime(worksheet, row, col, @datetime, format)

        /' Also write the format string for comparison. '/
        worksheet_write_string(worksheet, row, col + 1, date_formats(i), NULL)
    next i


    return workbook_close(workbook)
End Function

main()