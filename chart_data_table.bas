/'
 * An example of creating Excel column charts with data tables using the
 * libxlsxwriter library.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

/'
 * Write some data to the worksheet.
 '/
sub write_worksheet_data(byref worksheet as lxw_worksheet ptr , byref bold as lxw_format ptr)
    Dim As integer row, col
    Dim As uinteger data_(0 to 5, 0 to 2) = { _
        /' Three columns of data. '/ _
        {2, 10, 30}, _
        {3, 40, 60}, _
        {4, 50, 70}, _
        {5, 20, 50}, _
        {6, 10, 40}, _
        {7, 50, 30} _
    }

    worksheet_write_string(worksheet, CELL("A1"), "Number",  bold)
    worksheet_write_string(worksheet, CELL("B1"), "Batch 1", bold)
    worksheet_write_string(worksheet, CELL("C1"), "Batch 2", bold)

    for row = 0 to 5
        for col = 0 to 2
            worksheet_write_number(worksheet, row + 1, col, data_(row, col) , NULL)
        next
    next

end sub

/'
 * Create a worksheet with examples charts.
 '/
function main() as Integer

    var workbook  = new_workbook("chart_data_table.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)

    /' Add a bold format to use to highlight the header cells. '/
    var bold = workbook_add_format(workbook)
    format_set_bold(bold)

    /' Write some data for the chart. '/
    write_worksheet_data(worksheet, bold)


    /'
     * Chart 1. Create a column chart with a data table.
     '/
    var chart = workbook_add_chart(workbook, LXW_CHART_COLUMN)

    /' Add the first series to the chart. '/
    var series = chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7")

    /' Set the name for the series instead of the default "Series 1". '/
    chart_series_set_name(series, "=Sheet1!$B1$1")

    /' Add a second series but leave the categories and values undefined. They
     * can be defined later using the alternative syntax shown below.  '/
    series = chart_add_series(chart, NULL, NULL)

    /' Configure the series using a syntax that is easier to define programmatically. '/
    chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0) /' "=Sheet1!$A$2:$A$7" '/
    chart_series_set_values(series,     "Sheet1", 1, 2, 6, 2) /' "=Sheet1!$C$2:$C$7" '/
    chart_series_set_name_range(series, "Sheet1", 0, 2)       /' "=Sheet1!$C$1"      '/

    /' Add a chart title and some axis labels. '/
    chart_title_set_name(chart,        "Chart with Data Table")
    chart_axis_set_name(chart->x_axis, "Test number")
    chart_axis_set_name(chart->y_axis, "Sample length (mm)")

    /' Set a default data table on the X-axis. '/
    chart_set_table(chart)

    /' Insert the chart into the worksheet. '/
    worksheet_insert_chart(worksheet, CELL("E2"), chart)


    /'
     * Chart 2. Create a column chart with a data table and legend keys.
     '/
    chart = workbook_add_chart(workbook, LXW_CHART_COLUMN)

    /' Add the first series to the chart. '/
    series = chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7")

    /' Set the name for the series instead of the default "Series 1". '/
    chart_series_set_name(series, "=Sheet1!$B1$1")

    /' Add the second series to the chart. '/
    series = chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7")

    /' Set the name for the series instead of the default "Series 2". '/
    chart_series_set_name(series, "=Sheet1!$C1$1")

    /' Add a chart title and some axis labels. '/
    chart_title_set_name(chart,        "Data Table with legend keys")
    chart_axis_set_name(chart->x_axis, "Test number")
    chart_axis_set_name(chart->y_axis, "Sample length (mm)")

    /' Set a data table on the X-axis with the legend keys shown. '/
    chart_set_table(chart)
    chart_set_table_grid(chart, LXW_TRUE, LXW_TRUE, LXW_TRUE, LXW_TRUE)

    /' Turn off the legend. '/
    chart_legend_set_position(chart, LXW_CHART_LEGEND_NONE)


    /' Insert the chart into the worksheet. '/
    worksheet_insert_chart(worksheet, CELL("E18"), chart)


    return workbook_close(workbook)
End Function

main()