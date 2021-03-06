/'
 * An example of a clustered category chart using the libxlsxwriter library.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

/'
 * Write some data to the worksheet.
 '/
Sub write_worksheet_data(byref worksheet as lxw_worksheet ptr, byref bold as lxw_format ptr)

    worksheet_write_string(worksheet, 0, 0, "Types",      bold)
    worksheet_write_string(worksheet, 1, 0, "Type 1",     NULL)
    worksheet_write_string(worksheet, 4, 0, "Type 2",     NULL)

    worksheet_write_string(worksheet, 0, 1, "Sub Type",   bold)
    worksheet_write_string(worksheet, 1, 1, "Sub Type A", NULL)
    worksheet_write_string(worksheet, 2, 1, "Sub Type B", NULL)
    worksheet_write_string(worksheet, 3, 1, "Sub Type C", NULL)
    worksheet_write_string(worksheet, 4, 1, "Sub Type D", NULL)
    worksheet_write_string(worksheet, 5, 1, "Sub Type E", NULL)

    worksheet_write_string(worksheet, 0, 2, "Value 1",    bold)
    worksheet_write_number(worksheet, 1, 2, 5000,         NULL)
    worksheet_write_number(worksheet, 2, 2, 2000,         NULL)
    worksheet_write_number(worksheet, 3, 2, 250,          NULL)
    worksheet_write_number(worksheet, 4, 2, 6000,         NULL)
    worksheet_write_number(worksheet, 5, 2, 500,          NULL)

    worksheet_write_string(worksheet, 0, 3, "Value 2",    bold)
    worksheet_write_number(worksheet, 1, 3, 8000,         NULL)
    worksheet_write_number(worksheet, 2, 3, 3000,         NULL)
    worksheet_write_number(worksheet, 3, 3, 1000,         NULL)
    worksheet_write_number(worksheet, 4, 3, 6000,         NULL)
    worksheet_write_number(worksheet, 5, 3, 300,          NULL)

    worksheet_write_string(worksheet, 0, 4, "Value 3",    bold)
    worksheet_write_number(worksheet, 1, 4, 6000,         NULL)
    worksheet_write_number(worksheet, 2, 4, 4000,         NULL)
    worksheet_write_number(worksheet, 3, 4, 2000,         NULL)
    worksheet_write_number(worksheet, 4, 4, 6500,         NULL)
    worksheet_write_number(worksheet, 5, 4, 200,          NULL)

End Sub


/'
 * Create a worksheet with examples charts.
 '/
function main() as Integer

    var workbook  = new_workbook("chart_clustered2.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)
    var chart     = workbook_add_chart(workbook, LXW_CHART_COLUMN)

    /' Add a bold format to use to highlight the header cells. '/
    var bold = workbook_add_format(workbook)
    format_set_bold(bold)

    /' Write some data for the chart. '/
    write_worksheet_data(worksheet, bold)

    /'
     * Configure the series. Note, that the categories are 2D ranges (from
     * column A to column B). This creates the clusters. The series are shown
     * as formula strings for clarity but you can also use variables with the
     * chart_series_set_categories() and chart_series_set_values()
     * functions. See the docs.
     '/
    chart_add_series(chart, _
                     "=Sheet1!$A$2:$B$6", _
                     "=Sheet1!$C$2:$C$6")

    chart_add_series(chart, _
                     "=Sheet1!$A$2:$B$6", _
                     "=Sheet1!$D$2:$D$6")

    chart_add_series(chart, _
                     "=Sheet1!$A$2:$B$6", _
                     "=Sheet1!$E$2:$E$6")

    /' Set an Excel chart style. '/
    chart_set_style(chart, 37)

    /' Turn off the legend. '/
    chart_legend_set_position(chart, LXW_CHART_LEGEND_NONE)

    /' Insert the chart into the worksheet. '/
    worksheet_insert_chart(worksheet, CELL("G3"), chart)

    return workbook_close(workbook)
End Function

main()