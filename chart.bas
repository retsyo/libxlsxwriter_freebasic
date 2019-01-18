/'
 * An example of a simple Excel chart using the libxlsxwriter library.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

Rem Write some data to the worksheet. */
Sub write_worksheet_data(byref worksheet As lxw_worksheet Ptr)
    Rem Three columns of data. */ _
    Dim as Integer data_(0 to 4, 0 to 2) => { _
        {1,   2,   3}, _
        {2,   4,   6}, _
        {3,   6,   9}, _
        {4,   8,  12}, _
        {5,  10,  15} _
    }

    Dim as integer row, col
    for row = 0 to 4
        for col = 0 to 2
            worksheet_write_number(worksheet, row, col, data_(row, col), NULL)
            ' ~ worksheet_write_number(worksheet, row, col, 1, NULL)
        next
    next

End Sub

Rem Create a worksheet with a chart. */
function main() as Integer

    Dim as lxw_workbook ptr workbook  = new_workbook("chart.xlsx")
    Dim as lxw_worksheet ptr worksheet = workbook_add_worksheet(workbook, NULL)

    Rem Write some data for the chart. */
    write_worksheet_data(worksheet)

    Rem Create a chart object. */
    Dim as lxw_chart ptr chart = workbook_add_chart(workbook, LXW_CHART_COLUMN)

    /' Configure the chart. In simplest case we just add some value data
     * series. The NULL categories will default to 1 to 5 like in Excel.
     '/
    chart_add_series(chart, NULL, "Sheet1!$A$1:$A$5")
    chart_add_series(chart, NULL, "Sheet1!$B$1:$B$5")
    chart_add_series(chart, NULL, "Sheet1!$C$1:$C$5")

    Rem Insert the chart into the worksheet. */
    worksheet_insert_chart(worksheet, CELL("B7"), chart)

    return workbook_close(workbook)
End Function

main()
