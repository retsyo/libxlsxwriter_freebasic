/'
 * Example of how to create defined names using libxlsxwriter. This method is
 * used to define a user friendly name to represent a value, a single cell or
 * a range of cells in a workbook.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"


function main() as Integer

    Dim as lxw_workbook  ptr workbook = workbook_new("defined_name.xlsx")
    Dim as lxw_worksheet ptr worksheet

    /' We don't use the returned worksheets in this example and use a generic
     * loop instead. '/
    workbook_add_worksheet(workbook, NULL)
    workbook_add_worksheet(workbook, NULL)

    /' Define some global/workbook names. '/
    workbook_define_name(workbook, "Sales", "=!G1:H10")

    workbook_define_name(workbook, "Exchange_rate", "=0.96")
    workbook_define_name(workbook, "Sales",         "=Sheet1!$G$1:$H$10")

    /' Define a local/worksheet name. '/
    workbook_define_name(workbook, "Sheet2!Sales",  "=Sheet2!$G$1:$G$10")

    /' Write some text to the worksheets and one of the defined name in a formula. '/
    LXW_FOREACH_WORKSHEET(worksheet, workbook)
        worksheet_set_column(worksheet, 0, 0, 45, NULL)

        worksheet_write_string(worksheet, 0, 0, _
                               "This worksheet contains some defined names.", NULL)

        worksheet_write_string(worksheet, 1, 0, _
                               "See Formulas -> Name Manager above.", NULL)

        worksheet_write_string(worksheet, 2, 0, _
                               "Example formula in cell B3 ->", NULL)

        worksheet_write_formula(worksheet, 2, 1, "=Exchange_rate", NULL)


    return workbook_close(workbook)
End Function

main()