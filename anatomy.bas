/'
 * Anatomy of a simple libxlsxwriter program.
 *
 * Copyright 2014-2017, John McNamara, jmcnamara@cpan.org
 *
 * translated by Lee June by using https://github.com/retsyo/libxlsxwriter_freebasic
 '/

#include "auto_xlsxwriter.bi"

function main() as Integer

    Rem  Create a new workbook. */
    dim as lxw_workbook  ptr workbook   => workbook_new("anatomy.xlsx")

    Rem  Add a worksheet with a user defined sheet name. */
    dim as lxw_worksheet ptr worksheet1 => workbook_add_worksheet(workbook, "Demo")

    Rem  Add a worksheet with Excel's default sheet name: Sheet2. */
    dim as lxw_worksheet ptr worksheet2 => workbook_add_worksheet(workbook, NULL)

    Rem  Add some cell formats. */
    dim as lxw_format  ptr myformat1    => workbook_add_format(workbook)
    dim as lxw_format  ptr myformat2    = workbook_add_format(workbook)

    Rem  Set the bold property for the first format. */
    format_set_bold(myformat1)

    Rem  Set a number format for the second format. */
    format_set_num_format(myformat2, "$#,##0.00")

    Rem  Widen the first column to make the text clearer. */
    worksheet_set_column(worksheet1, 0, 0, 20, NULL)

    Rem  Write some unformatted data. */
    worksheet_write_string(worksheet1, 0, 0, "Peach", NULL)
    worksheet_write_string(worksheet1, 1, 0, "Plum",  NULL)

    Rem  Write formatted data. */
    worksheet_write_string(worksheet1, 2, 0, "Pear",  myformat1)

    Rem  Formats can be reused. */
    worksheet_write_string(worksheet1, 3, 0, "Persimmon",  myformat1)


    Rem  Write some numbers. */
    worksheet_write_number(worksheet1, 5, 0, 123,       NULL)
    worksheet_write_number(worksheet1, 6, 0, 4567.555,  myformat2)


    Rem  Write to the second worksheet. */
    worksheet_write_string(worksheet2, 0, 0, "Some text", myformat1)


    Rem  Close the workbook, save the file and free any memory. */
   dim as lxw_error errcode => workbook_close(workbook)

    Rem  Check if there was any error creating the xlsx file. */
    if errcode then
        printf( !"Error in workbook_close().\nError %d = %s\n", _
            errcode, lxw_strerror(errcode))
    end if

    return errcode

end Function

main()
