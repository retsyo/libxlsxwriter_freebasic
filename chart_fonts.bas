/'
 * An example of a simple Excel chart with user defined fonts using the
 * libxlsxwriter library.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"


/' Create a worksheet with a chart. '/
function main() as Integer

    var workbook  = new_workbook("chart_fonts.xlsx")
    var worksheet = workbook_add_worksheet(workbook, NULL)

    /' Write some data for the chart. '/
    worksheet_write_number(worksheet, 0, 0, 10, NULL)
    worksheet_write_number(worksheet, 1, 0, 40, NULL)
    worksheet_write_number(worksheet, 2, 0, 50, NULL)
    worksheet_write_number(worksheet, 3, 0, 20, NULL)
    worksheet_write_number(worksheet, 4, 0, 10, NULL)
    worksheet_write_number(worksheet, 5, 0, 50, NULL)

    /' Create a chart object. '/
    var chart = workbook_add_chart(workbook, LXW_CHART_LINE)

    /' Configure the chart. '/
    chart_add_series(chart, NULL, "Sheet1!$A$1:$A$6")

    /' Create some fonts to use in the chart.  '/
    Dim As lxw_chart_font font1
    With font1
        .name = StrPtr("Calibri")
        .color = LXW_COLOR_BLUE
    End With

    Dim As lxw_chart_font font2
    With font2
        .name = @"Courier"
        .color = &H92D050
    End With

    Dim As lxw_chart_font font3
    With font3
        .name = @"Arial"
        .color = &H00B0F0
    End With

    Dim As lxw_chart_font font4
    With font4
        .name = @"Century"
        .color = LXW_COLOR_RED
    End With

    Dim As lxw_chart_font font5
    With font5
        .rotation = -30
    End With

    Dim As lxw_chart_font font6
    With font6
        .bold      = LXW_TRUE
        .italic    = LXW_TRUE
        .underline = LXW_TRUE
        .color     = &H7030A0
    End With

    /' Write the chart title with a font. '/
    chart_title_set_name(chart, "Test Results")
    chart_title_set_name_font(chart, @font1)

    /' Write the Y axis with a font. '/
    chart_axis_set_name(chart->y_axis, "Units")
    chart_axis_set_name_font(chart->y_axis, @font2)
    chart_axis_set_num_font(chart->y_axis, @font3)

    /' Write the X axis with a font. '/
    chart_axis_set_name(chart->x_axis, "Month")
    chart_axis_set_name_font(chart->x_axis, @font4)
    chart_axis_set_num_font(chart->x_axis, @font5)

    /' Display the chart legend at the bottom of the chart. '/
    chart_legend_set_position(chart, LXW_CHART_LEGEND_BOTTOM)
    chart_legend_set_font(chart, @font6)

    /' Insert the chart into the worksheet. '/
    worksheet_insert_chart(worksheet, CELL("C1"), chart)

    return workbook_close(workbook)
End Function

main()