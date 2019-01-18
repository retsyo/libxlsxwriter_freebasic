/'
 * Example of adding an autofilter to a worksheet in Excel using
 * libxlsxwriter.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"


function main() as Integer

    Dim As lxw_workbook Ptr workbook  = workbook_new("autofilter.xlsx")
    Dim As lxw_worksheet Ptr worksheet = workbook_add_worksheet(workbook, NULL)
    Dim As Integer i


    Rem Simple data structure to represent the row data. */
    Type row
        Dim as String * 16 region
        Dim as String * 16 item
        Dim As Integer volume
        Dim as String * 16 Month
    End Type

    Dim As row data_(0 to 49) => {    _
        ("East",  "Apple",   9000, "July"      ),    _
        ("East",  "Apple",   5000, "July"      ),    _
        ("South", "Orange",  9000, "September" ),    _
        ("North", "Apple",   2000, "November"  ),    _
        ("West",  "Apple",   9000, "November"  ),    _
        ("South", "Pear",    7000, "October"   ),    _
        ("North", "Pear",    9000, "August"    ),    _
        ("West",  "Orange",  1000, "December"  ),    _
        ("West",  "Grape",   1000, "November"  ),    _
        ("South", "Pear",   10000, "April"     ),    _
        ("West",  "Grape",   6000, "January"   ),    _
        ("South", "Orange",  3000, "May"       ),    _
        ("North", "Apple",   3000, "December"  ),    _
        ("South", "Apple",   7000, "February"  ),    _
        ("West",  "Grape",   1000, "December"  ),    _
        ("East",  "Grape",   8000, "February"  ),    _
        ("South", "Grape",  10000, "June"      ),    _
        ("West",  "Pear",    7000, "December"  ),    _
        ("South", "Apple",   2000, "October"   ),    _
        ("East",  "Grape",   7000, "December"  ),    _
        ("North", "Grape",   6000, "April"     ),    _
        ("East",  "Pear",    8000, "February"  ),    _
        ("North", "Apple",   7000, "August"    ),    _
        ("North", "Orange",  7000, "July"      ),    _
        ("North", "Apple",   6000, "June"      ),    _
        ("South", "Grape",   8000, "September" ),    _
        ("West",  "Apple",   3000, "October"   ),    _
        ("South", "Orange", 10000, "November"  ),    _
        ("West",  "Grape",   4000, "July"      ),    _
        ("North", "Orange",  5000, "August"    ),    _
        ("East",  "Orange",  1000, "November"  ),    _
        ("East",  "Orange",  4000, "October"   ),    _
        ("North", "Grape",   5000, "August"    ),    _
        ("East",  "Apple",   1000, "December"  ),    _
        ("South", "Apple",   10000, "March"    ),    _
        ("East",  "Grape",   7000, "October"   ),    _
        ("West",  "Grape",   1000, "September" ),    _
        ("East",  "Grape",  10000, "October"   ),    _
        ("South", "Orange",  8000, "March"     ),    _
        ("North", "Apple",   4000, "July"      ),    _
        ("South", "Orange",  5000, "July"      ),    _
        ("West",  "Apple",   4000, "June"      ),    _
        ("East",  "Apple",   5000, "April"     ),    _
        ("North", "Pear",    3000, "August"    ),    _
        ("East",  "Grape",   9000, "November"  ),    _
        ("North", "Orange",  8000, "October"   ),    _
        ("East",  "Apple",  10000, "June"      ),    _
        ("South", "Pear",    1000, "December"  ),    _
        ("North", "Grape",   10000, "July"     ),    _
        ("East",  "Grape",   6000, "February"  )     _
    }

    Rem Write the column headers. */
    worksheet_write_string(worksheet, 0, 0, "Region", NULL)
    worksheet_write_string(worksheet, 0, 1, "Item",   NULL)
    worksheet_write_string(worksheet, 0, 2, "Volume" , NULL)
    worksheet_write_string(worksheet, 0, 3, "Month",  NULL)


    print sizeof(data_), sizeof(row), sizeof(data_)/sizeof(row)
    print len(data_), len(row), sizeof(data_)/sizeof(row)
    Rem Write the row data. */
    for i = 0 to Ubound(data_) - Lbound(data_) 'sizeof(data_)/sizeof(row)
        worksheet_write_string(worksheet, i + 1, 0, data_(i).region, NULL)
        worksheet_write_string(worksheet, i + 1, 1, data_(i).item,   NULL)
        worksheet_write_number(worksheet, i + 1, 2, data_(i).volume , NULL)
        worksheet_write_string(worksheet, i + 1, 3, data_(i).month,  NULL)
    Next


    Rem Add the autofilter. */
    worksheet_autofilter(worksheet, 0, 0, 50, 3)

    return workbook_close(workbook)
End Function

main()