Rem written by Lee June

#define NULL 0

Type lxw_workbook as integer
Type lxw_worksheet as integer
Type lxw_format as integer

declare function workbook_new alias "workbook_new" (byval filename as const zstring ptr) as  any ptr

declare function workbook_close alias "workbook_close"(byval workbook as any ptr) as Integer

declare function workbook_add_worksheet alias "workbook_add_worksheet"(byval workbook as any ptr, byval sheetname as const zstring ptr=0) as Integer ptr

declare sub format_set_bold alias "format_set_bold" (byval fmt as lxw_format ptr)

declare Function workbook_add_format alias "workbook_add_format" (byval workbook as any ptr) as lxw_format ptr

declare function worksheet_set_column alias "worksheet_set_column" (byval worksheet as lxw_worksheet ptr, byval first_col as ushort, byval last_col as ushort, byval width as double, byval format as lxw_format ptr) as long

declare function worksheet_write_string alias "worksheet_write_string"(byval worksheet as lxw_worksheet ptr, byval row as ushort, byval col as ushort, byval string as const zstring ptr, byval format as lxw_format ptr=NULL) as long


declare function worksheet_write_number alias "worksheet_write_number" (byval worksheet as lxw_worksheet ptr, byval row as ushort, byval col as ushort, byval number as double, byval format as lxw_format ptr) as long

declare function worksheet_insert_image alias "worksheet_insert_image"(byval worksheet as lxw_worksheet ptr, byval row as ushort, byval col as ushort, byval filename as const zstring ptr) as long

function main() as integer

    /' Create a new workbook and add a worksheet. '/
    dim as lxw_workbook  ptr workbook  = workbook_new("test.xlsx")
    dim as lxw_worksheet ptr worksheet = workbook_add_worksheet(workbook, NULL)

    /' Add a format. '/
    dim as lxw_format ptr format = workbook_add_format(workbook)

    /' Set the bold property for the format '/
    format_set_bold(format)

    /' Change the column width for clarity. '/
    worksheet_set_column(worksheet, 0, 0, 20, NULL)

    /' Write some simple text. '/
    worksheet_write_string(worksheet, 0, 0, "Hello", NULL)

    ' ~ /' Text with formatting. '/
    worksheet_write_string(worksheet, 1, 0, "World", format)

    ' ~ /' Write some numbers. '/
    worksheet_write_number(worksheet, 2, 0, 123,     NULL)
    worksheet_write_number(worksheet, 3, 0, 123.456, NULL)

    ' ~ /' Insert an image. '/
    ' ~ worksheet_insert_image(worksheet, 1, 2, "logo.png")

    workbook_close(workbook)

    return 0
end function


main()