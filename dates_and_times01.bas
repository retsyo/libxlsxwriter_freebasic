/'
 * Example of writing a dates and time in Excel using a number with date
 * formatting. This demonstrates that dates and times in Excel are just
 * formatted real numbers.
 *
 * An easier approach using a lxw_datetime struct is shown in example
 * dates_and_times02.c.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
'/

#include "auto_xlsxwriter.bi"

function main() as Integer

    /' A number to display as a date. '/
    Dim as double number = 41333.5

    /' Create a new workbook and add a worksheet. '/
    dim as lxw_workbook  ptr workbook  = workbook_new("dates_and_times01.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)

    /' Add a format with date formatting. '/
    dim as lxw_format    ptr format    = workbook_add_format(workbook)
    format_set_num_format(format, "mmm d yyyy hh:mm AM/PM")

    /' Widen the first column to make the text clearer. '/
    worksheet_set_column(worksheet, 0, 0, 20, NULL)

    /' Write the number without formatting. '/
    worksheet_write_number(worksheet, 0, 0, number, NULL   )  Rem 41333.5

    /' Write the number with formatting. Note: the worksheet_write_datetime()
     * function is preferable for writing dates and times. This is for
     * demonstration purposes only.
     '/
    worksheet_write_number(worksheet, 1, 0, number, format)   Rem Feb 28 2013 12:00 PM

    return workbook_close(workbook)
End Function

main()