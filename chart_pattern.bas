/'
 * An example of a simple Excel chart with patterns using the libxlsxwriter
 * library.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"


/' Create a worksheet with a chart. '/
function main() as Integer

    var workbook  = new_workbook("chart_pattern.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)


    /' Add a bold format to use to highlight the header cells. '/
    var bold = workbook_add_format(workbook)
    format_set_bold(bold)

    /' Write some data for the chart. '/
    worksheet_write_string(worksheet, 0, 0, "Shingle", bold)
    worksheet_write_number(worksheet, 1, 0, 105,       NULL)
    worksheet_write_number(worksheet, 2, 0, 150,       NULL)
    worksheet_write_number(worksheet, 3, 0, 130,       NULL)
    worksheet_write_number(worksheet, 4, 0, 90,        NULL)
    worksheet_write_string(worksheet, 0, 1, "Brick",   bold)
    worksheet_write_number(worksheet, 1, 1, 50,        NULL)
    worksheet_write_number(worksheet, 2, 1, 120,       NULL)
    worksheet_write_number(worksheet, 3, 1, 100,       NULL)
    worksheet_write_number(worksheet, 4, 1, 110,       NULL)

    /' Create a chart object. '/
    var chart = workbook_add_chart(workbook, LXW_CHART_COLUMN)

    /' Configure the chart. '/
    var series1 = chart_add_series(chart, NULL, "Sheet1!$A$2:$A$5")
    var series2 = chart_add_series(chart, NULL, "Sheet1!$B$2:$B$5")

    chart_series_set_name(series1, "=Sheet1!$A1$1")
    chart_series_set_name(series2, "=Sheet1!$B1$1")

    chart_title_set_name(chart,        "Cladding types")
    chart_axis_set_name(chart->x_axis, "Region")
    chart_axis_set_name(chart->y_axis, "Number of houses")


    /' Configure an add the chart series patterns. '/
    Dim As lxw_chart_pattern pattern1
    With pattern1
        .type = LXW_CHART_PATTERN_SHINGLE
        .fg_color = &H804000
        .bg_color = &HC68C53
    End With

    Dim As lxw_chart_pattern pattern2
    With pattern2
        .type = LXW_CHART_PATTERN_HORIZONTAL_BRICK
        .fg_color = &HB30000
        .bg_color = &HFF6666
    End With

    chart_series_set_pattern(series1, @pattern1)
    chart_series_set_pattern(series2, @pattern2)

    /' Configure and set the chart series borders. '/
    Dim As lxw_chart_line line1
    line1.color = &H804000

    Dim As lxw_chart_line line2
    line2.color = &Hb30000

    chart_series_set_line(series1, @line1)
    chart_series_set_line(series2, @line2)

    /' Widen the gap between the series/categories. '/
    chart_set_series_gap(chart, 70)

    /' Insert the chart into the worksheet. '/
    worksheet_insert_chart(worksheet, CELL("D2"), chart)

    return workbook_close(workbook)
End Function

main()