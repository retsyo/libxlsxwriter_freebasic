/'
 * An example of a simple Excel chart using the libxlsxwriter library. This
 * example is used in the "Working with Charts" section of the docs.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"


/' Create a worksheet with a chart. '/
function main() as Integer

    Dim As lxw_workbook  ptr workbook  = new_workbook("chart_line.xlsx")
    Dim As lxw_worksheet ptr worksheet = workbook_add_worksheet(workbook, NULL)
    Dim As lxw_chart ptr chart
    Dim As lxw_chart_series ptr series

    /' Write some data for the chart. '/
    worksheet_write_number(worksheet, 0, 0, 10, NULL)
    worksheet_write_number(worksheet, 1, 0, 40, NULL)
    worksheet_write_number(worksheet, 2, 0, 50, NULL)
    worksheet_write_number(worksheet, 3, 0, 20, NULL)
    worksheet_write_number(worksheet, 4, 0, 10, NULL)
    worksheet_write_number(worksheet, 5, 0, 50, NULL)

    /' Create a chart object. '/
    chart = workbook_add_chart(workbook, LXW_CHART_LINE)

    /' Configure the chart. '/
    series = chart_add_series(chart, NULL, "Sheet1!$A$1:$A$6")

    ' ~ (void)series; /' Do something with series in the real examples. '/

    /' Insert the chart into the worksheet. '/
    worksheet_insert_chart(worksheet, CELL("C1"), chart)

    return workbook_close(workbook)
End Function

main()
