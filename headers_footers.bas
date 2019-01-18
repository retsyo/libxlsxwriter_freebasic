/'
 * This program shows several examples of how to set up headers and
 * footers with libxlsxwriter.
 *
 * The control characters used in the header/footer strings are:
 *
 *     Control             Category            Description
 *     =======             ========            ===========
 *     &L                  Justification       Left
 *     &C                                      Center
 *     &R                                      Right
 *     &P                  Information         Page number
 *     &N                                      Total number of pages
 *     &D                                      Date
 *     &T                                      Time
 *     &F                                      File name
 *     &A                                      Worksheet name
 *     &fontsize           Font                Font size
 *     &"font,style"                           Font name and style
 *     &U                                      Single underline
 *     &E                                      Double underline
 *     &S                                      Strikethrough
 *     &X                                      Superscript
 *     &Y                                      Subscript
 *     &[Picture]          Images              Image placeholder
 *     &G                                      Same as &[Picture]
 *     &&                  Miscellaneous       Literal ampersand &
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

function main() as Integer

    Dim as lxw_workbook ptr workbook  = workbook_new("headers_footers.xlsx")

    Dim as String preview = "Select Print Preview to see the header and footer"

    /'
     * A simple example to start
     '/
    Dim as lxw_worksheet Ptr worksheet1 = workbook_add_worksheet(workbook, "Simple")
    Dim As String header1 = "&CHere is some centered text."
    Dim As String footer1 = "&LHere is some left aligned text."

    worksheet_set_header(worksheet1, header1)
    worksheet_set_footer(worksheet1, footer1)

    worksheet_set_column(worksheet1, 0, 0, 50, NULL)
    worksheet_write_string(worksheet1, 0, 0, preview, NULL)


    /'
     * This is an example of some of the header/footer variables.
     '/
    Dim as lxw_worksheet Ptr worksheet2 = workbook_add_worksheet(workbook, "Variables")
    Dim AS String header2 = "&LPage &P of &N" "&CFilename: &F" "&RSheetname: &A"
    Dim AS String footer2 = "&LCurrent date: &D" "&RCurrent time: &T"
    Dim As lxw_row_t breaks(0 to 1) = {20, 0}

    worksheet_set_header(worksheet2, header2)
    worksheet_set_footer(worksheet2, footer2)

    worksheet_set_column(worksheet2, 0, 0, 50, NULL)
    worksheet_write_string(worksheet2, 0, 0, preview, NULL)

    ' ~ worksheet_set_h_pagebreaks(worksheet2, breaks)
    worksheet_write_string(worksheet2, 20, 0, "Next page", NULL)


    /'
     * This example shows how to use more than one font.
     '/
    Dim as lxw_worksheet Ptr worksheet3 = workbook_add_worksheet(workbook, "Mixed fonts")
    Dim AS String header3 = !"&C&\"Courier New,Bold\"Hello &\"Arial,Italic\"World"
    Dim AS String footer3 = !"&C&\"Symbol\"e&\"Arial\" = mc&X2"

    worksheet_set_header(worksheet3, header3)
    worksheet_set_footer(worksheet3, footer3)

    worksheet_set_column(worksheet3, 0, 0, 50, NULL)
    worksheet_write_string(worksheet3, 0, 0, preview, NULL)


    /'
     * Example of line wrapping.
     '/
    Dim As lxw_worksheet ptr worksheet4 = workbook_add_worksheet(workbook, "Word wrap")
    Dim AS String header4 = !"&CHeading 1\nHeading 2"

    worksheet_set_header(worksheet4, header4)

    worksheet_set_column(worksheet4, 0, 0, 50, NULL)
    worksheet_write_string(worksheet4, 0, 0, preview, NULL)


    /'
     * Example of inserting a literal ampersand &
     '/
    Dim AS lxw_worksheet ptr worksheet5 = workbook_add_worksheet(workbook, "Ampersand")
    Dim AS String header5 = "&CCuriouser && Curiouser - Attorneys at Law"

    worksheet_set_header(worksheet5, header5)

    worksheet_set_column(worksheet5, 0, 0, 50, NULL)
    worksheet_write_string(worksheet5, 0, 0, preview, NULL)


    workbook_close(workbook)

    return 0
End Function

main()