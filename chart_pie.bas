/'
 * An example of creating an Excel pie chart using the libxlsxwriter library.
 *
 * The demo also shows how to set segment colors. It is possible to define
 * chart colors for most types of libxlsxwriter charts via the series
 * formatting functions. However, Pie/Doughnut charts are a special case since
 * each segment is represented as a point so it is necessary to assign
 * formatting to each point in the series.
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

    worksheet_write_string(worksheet, CELL("A1"), "Category", bold)
    worksheet_write_string(worksheet, CELL("A2"), "Apple",    NULL)
    worksheet_write_string(worksheet, CELL("A3"), "Cherry",   NULL)
    worksheet_write_string(worksheet, CELL("A4"), "Pecan",    NULL)

    worksheet_write_string(worksheet, CELL("B1"), "Values",   bold)
    worksheet_write_number(worksheet, CELL("B2"), 60,         NULL)
    worksheet_write_number(worksheet, CELL("B3"), 30,         NULL)
    worksheet_write_number(worksheet, CELL("B4"), 10,         NULL)
End Sub

/'
 * Create a worksheet with examples charts.
 *
 '/
function main() as Integer

    var workbook  = new_workbook("chart_pie.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)

    /' Add a bold format to use to highlight the header cells. '/
    var bold = workbook_add_format(workbook)
    format_set_bold(bold)

    /' Write some data for the chart. '/
    write_worksheet_data(worksheet, bold)


    /'
     * Chart 1: Create a simple pie chart.
     '/
    var chart = workbook_add_chart(workbook, LXW_CHART_PIE)

    /' Add the first series to the chart. '/
    var series = chart_add_series(chart, "=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4")

    /' Set the name for the series instead of the default "Series 1". '/
    chart_series_set_name(series, "Pie sales data")

    /' Add a chart title. '/
    chart_title_set_name(chart, "Popular Pie Types")

    /' Set an Excel chart style. '/
    chart_set_style(chart, 10)

    /' Insert the chart into the worksheet. '/
    worksheet_insert_chart(worksheet, CELL("D2"), chart)


    /'
     * Chart 2: Create a pie chart with user defined segment colors.
     '/
    chart = workbook_add_chart(workbook, LXW_CHART_PIE)

    /' Add the first series to the chart. '/
    series = chart_add_series(chart, "=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4")

    /' Set the name for the series instead of the default "Series 1". '/
    chart_series_set_name(series, "Pie sales data")

    /' Add a chart title. '/
    chart_title_set_name(chart, "Pie Chart with user defined colors")

    /' Add for fills for use in the chart. '/
    Dim As lxw_chart_fill fill1
    fill1.color = &H5ABA10

    Dim As lxw_chart_fill fill2
    fill2.color = &HFE110E

    Dim As lxw_chart_fill fill3
    fill3.color = &HCA5C05


    /' Add some points with the above fills. '/
    Dim As lxw_chart_point point1
    point1.fill = @fill1

    Dim As lxw_chart_point point2
    point2.fill = @fill2

    Dim As lxw_chart_point point3
    point3.fill = @fill3


    /' Create an array of the point objects. '/
    Dim As lxw_chart_point ptr points(0 to 3) = { @point1, _
                                 @point2, _
                                 @point3, _
                                 NULL}


    /' Add/override the points/segments of the chart. '/
    chart_series_set_points(series, @points(0))

    /' Insert the chart into the worksheet. '/
    worksheet_insert_chart(worksheet, CELL("D18"), chart)


    /'
     * Chart 3: Create a pie chart with rotation of the segments.
     '/
    chart = workbook_add_chart(workbook, LXW_CHART_PIE)

    /' Add the first series to the chart. '/
    series = chart_add_series(chart, "=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4")

    /' Set the name for the series instead of the default "Series 1". '/
    chart_series_set_name(series, "Pie sales data")

    /' Add a chart title. '/
    chart_title_set_name(chart, "Pie Chart with segment rotation")


    /' Change the angle/rotation of the first segment. '/
    chart_set_rotation(chart, 90)

    /' Insert the chart into the worksheet. '/
    worksheet_insert_chart(worksheet, CELL("D34"), chart)


    return workbook_close(workbook)
End Function

main()