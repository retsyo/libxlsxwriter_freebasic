/'
 * Example of setting document properties such as Author, Title, etc., for an
 * Excel spreadsheet using libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

function main() as Integer

    Dim as lxw_workbook  ptr workbook   = workbook_new("doc_properties.xlsx")
    Dim as lxw_worksheet ptr worksheet  = workbook_add_worksheet(workbook, NULL)

    /' Create a properties structure and set some of the fields. '/
    Dim as lxw_doc_properties properties
    With properties
        .title    = @"This is an example spreadsheet"
        .subject  = @"With document properties"
        .author   = @"John McNamara"
        .manager  = @"Dr. Heinz Doofenshmirtz"
        .company  = @"of Wolves"
        .category = @"Example spreadsheets"
        .keywords = @"Sample, Example, Properties"
        .comments = @"Created with libxlsxwriter"
        .status   = @"Quo"
    End With

    /' Set the properties in the workbook. '/
    workbook_set_properties(workbook, @properties)

    /' Add some text to the file. '/
    worksheet_set_column(worksheet, 0, 0, 50, NULL)
    worksheet_write_string(worksheet, 0, 0, _
                           "Select 'Workbook Properties' to see properties." , NULL)

    workbook_close(workbook)

    return 0
End Function

main()