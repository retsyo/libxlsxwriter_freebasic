/'
 * An example showing all 48 default chart styles available in Excel 2007
 * using the libxlsxwriter library. Note, these styles are not the same as the
 * styles available in Excel 2013.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"


function main() as Integer

    Dim As Long chart_types(0 to 4)   => {LXW_CHART_COLUMN, LXW_CHART_AREA, LXW_CHART_LINE, LXW_CHART_PIE}

    Dim As String chart_names(0 to 4)  => {"Column", "Area", "Line", "Pie"}

    Dim As string *32 chart_title

    Dim As Long row_num, col_num, chart_num, style_num
    dim as lxw_worksheet ptr worksheet
    Dim As lxw_chart ptr chart

    var workbook  = new_workbook("chart_styles.xlsx")

    for chart_num = 0 to 3
        /' Add a worksheet for each chart type. '/
        worksheet = workbook_add_worksheet(workbook, chart_names(chart_num))
        worksheet_set_zoom(worksheet, 30)

        /' Create 48 charts, each with a different style. '/
        style_num = 1
        for row_num = 0  to 89 step 15
            for col_num = 0 to 63 step 8
                chart = workbook_add_chart(workbook, chart_types(chart_num))
                sprintf(chart_title, "Style %d", style_num)

                chart_add_series(chart, NULL, "=Data!$A$1:$A$6")
                chart_title_set_name(chart, chart_title)
                chart_set_style(chart, style_num)

                worksheet_insert_chart(worksheet, row_num, col_num, chart)

                style_num = 1 + style_num
            Next

        Next
    Next

    /' Create a worksheet with data for the charts. '/
    worksheet = workbook_add_worksheet(workbook, "Data")
    worksheet_write_number(worksheet, 0, 0, 10, NULL)
    worksheet_write_number(worksheet, 1, 0, 40, NULL)
    worksheet_write_number(worksheet, 2, 0, 50, NULL)
    worksheet_write_number(worksheet, 3, 0, 20, NULL)
    worksheet_write_number(worksheet, 4, 0, 10, NULL)
    worksheet_write_number(worksheet, 5, 0, 50, NULL)

    return workbook_close(workbook)
End Function

main()