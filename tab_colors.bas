/'
 * Example of how to set Excel worksheet tab colors using libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

function main() as Integer

    Dim as lxw_workbook  ptr workbook   = workbook_new("tab_colors.xlsx")

    /' Set up some worksheets. '/
    Dim As lxw_worksheet ptr worksheet1 = workbook_add_worksheet(workbook, NULL)
    Dim As lxw_worksheet ptr worksheet2 = workbook_add_worksheet(workbook, NULL)
    Dim As lxw_worksheet ptr worksheet3 = workbook_add_worksheet(workbook, NULL)
    Dim As lxw_worksheet ptr worksheet4 = workbook_add_worksheet(workbook, NULL)


    /' Set the tab colors. '/
    worksheet_set_tab_color(worksheet1, LXW_COLOR_RED)
    worksheet_set_tab_color(worksheet2, LXW_COLOR_GREEN)
    worksheet_set_tab_color(worksheet3, &HFF9900) /' Orange. '/

    /' worksheet4 will have the default color. '/
    worksheet_write_string(worksheet4, 0, 0, "Hello", NULL)

    workbook_close(workbook)

    return 0
End Function

main()