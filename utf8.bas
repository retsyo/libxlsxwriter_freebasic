/'
 * A simple Unicode UTF-8 example using libxlsxwriter.
 *
 * Note: The source file must be UTF-8 encoded.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

function main() as Integer

    Dim as lxw_workbook ptr workbook  = workbook_new("utf8.xlsx")
    Dim as lxw_worksheet ptr worksheet = workbook_add_worksheet(workbook, NULL)

    worksheet_write_string(worksheet, 2, 1, "Это фраза на русском!", NULL)
    worksheet_write_string(worksheet, 3, 1, "汉字", NULL)

    return workbook_close(workbook)
End Function

main()
