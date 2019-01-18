/'
 * Example of setting custom document properties for an Excel spreadsheet
 * using libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

function main() as Integer

    var workbook   = workbook_new("doc_custom_properties.xlsx")
    var worksheet  = workbook_add_worksheet(workbook, NULL)
    Dim As lxw_datetime   datetime  = (2016, 12, 12,  0, 0, 0.0)

    /' Set some custom document properties in the workbook. '/
    workbook_set_custom_property_string  (workbook, "Checked by",      "Eve")
    workbook_set_custom_property_datetime(workbook, "Date completed",   @datetime)
    workbook_set_custom_property_number  (workbook, "Document number",  12345)
    workbook_set_custom_property_number  (workbook, "Reference number", 1.2345)
    workbook_set_custom_property_boolean (workbook, "Has Review",       1)
    workbook_set_custom_property_boolean (workbook, "Signed off",       0)


    /' Add some text to the file. '/
    worksheet_set_column(worksheet, 0, 0, 50, NULL)
    worksheet_write_string(worksheet, 0, 0, _
                           "Select 'Workbook Properties' to see properties." , NULL)

    workbook_close(workbook)

    return 0
End Function

main()