Rem written by Lee June

' ~ #include "zlib.bi"
' ~ #inclib "libxlsxwriter"
' ~ #inclib "libxlsxwriter_static"

#define NULL 0
declare function workbook_new alias "workbook_new" (byval filename as const zstring ptr) as  any ptr

declare function workbook_close alias "workbook_close"(byval workbook as any ptr) as Integer
declare function workbook_add_worksheet alias "workbook_add_worksheet"(byval workbook as any ptr, byval sheetname as const zstring ptr=0) as Integer ptr

declare sub worksheet_write_string alias "worksheet_write_string" (byval workbook as any ptr, _
    byval row as Integer, byval col as Integer, byval value as const zstring ptr, byval fmt as any ptr=NULL)

Type lxw_workbook as integer
Type lxw_worksheet as integer

Dim as Integer i

dim as lxw_workbook ptr workbook  => workbook_new("mini.xlsx")
dim as lxw_worksheet ptr worksheet => workbook_add_worksheet(workbook, "NULL")

worksheet_write_string(worksheet, 1, 3, "Hello Excel", NULL)

workbook_close(workbook)
