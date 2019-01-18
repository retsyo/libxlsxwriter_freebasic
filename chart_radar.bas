/'
 * An example of creating Excel radar charts using the libxlsxwriter library.
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
    Dim As uinteger data_(..., ...) = { _
        /' Three columns of data. '/ _
        {2, 30, 25}, _
        {3, 60, 40}, _
        {4, 70, 50}, _
        {5, 50, 30}, _
        {6, 40, 50}, _
        {7, 30, 40} _
    }

    worksheet_write_string(worksheet, CELL("A1"), "Number",  bold)
    worksheet_write_string(worksheet, CELL("B1"), "Batch 1", bold)
    worksheet_write_string(worksheet, CELL("C1"), "Batch 2", bold)

    for row = Lbound(data_, 1) to Ubound(data_, 1)
        for col = Lbound(data_, 2) to Ubound(data_, 2)
            worksheet_write_number(worksheet, row + 1, col, data_(row, col) , NULL)
        next
    next

end sub

/'
 * Create a worksheet with examples charts.
 '/
function main() as Integer

    var workbook  = new_workbook("chart_radar.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)

    /' Add a bold format to use to highlight the header cells. '/
    var bold = workbook_add_format(workbook)
    format_set_bold(bold)

    /' Write some data for the chart. '/
    write_worksheet_data(worksheet, bold)


    /'
     * Chart 1. Create a radar chart.
     '/
    var chart = workbook_add_chart(workbook, LXW_CHART_RADAR)

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
    chart_title_set_name(chart,        "Results of sample analysis")
    chart_axis_set_name(chart->x_axis, "Test number")
    chart_axis_set_name(chart->y_axis, "Sample length (mm)")

    /' Set an Excel chart style. '/
    chart_set_style(chart, 11)

    /' Insert the chart into the worksheet. '/
    worksheet_insert_chart(worksheet, CELL("E2"), chart)


    /'
     * Chart 2. Create a radar chart with markers.
     '/
    chart = workbook_add_chart(workbook, LXW_CHART_RADAR_WITH_MARKERS)

    /' Add the first series to the chart. '/
    series = chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7")

    /' Set the name for the series instead of the default "Series 1". '/
    chart_series_set_name(series, "=Sheet1!$B1$1")

    /' Add the second series to the chart. '/
    series = chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7")

    /' Set the name for the series instead of the default "Series 2". '/
    chart_series_set_name(series, "=Sheet1!$C1$1")

    /' Add a chart title and some axis labels. '/
    chart_title_set_name(chart,        "Results of sample analysis")
    chart_axis_set_name(chart->x_axis, "Test number")
    chart_axis_set_name(chart->y_axis, "Sample length (mm)")

    /' Set an Excel chart style. '/
    chart_set_style(chart, 12)

    /' Insert the chart into the worksheet. '/
    worksheet_insert_chart(worksheet, CELL("E18"), chart)


    /'
     * Chart 3. Create a filled radar chart.
     '/
    chart = workbook_add_chart(workbook, LXW_CHART_RADAR_FILLED)

    /' Add the first series to the chart. '/
    series = chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7")

    /' Set the name for the series instead of the default "Series 1". '/
    chart_series_set_name(series, "=Sheet1!$B1$1")

    /' Add the second series to the chart. '/
    series = chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7")

    /' Set the name for the series instead of the default "Series 2". '/
    chart_series_set_name(series, "=Sheet1!$C1$1")

    /' Add a chart title and some axis labels. '/
    chart_title_set_name(chart,        "Results of sample analysis")
    chart_axis_set_name(chart->x_axis, "Test number")
    chart_axis_set_name(chart->y_axis, "Sample length (mm)")

    /' Set an Excel chart style. '/
    chart_set_style(chart, 13)

    /' Insert the chart into the worksheet. '/
    worksheet_insert_chart(worksheet, CELL("E34"), chart)


    return workbook_close(workbook)
End Function

main()