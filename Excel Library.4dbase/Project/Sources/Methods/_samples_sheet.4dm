//%attributes = {"invisible":true}
//methods:
//http://msdn.microsoft.com/en-us/library/bb259445(v=office.12).aspx
//properties:
//http://msdn.microsoft.com/en-us/library/bb259443(v=office.12).aspx

XCEL_workbook_close_all
$workbook:=XCEL_workbook_create
XCEL_application_show

$count_l:=XCEL_sheet_count($workbook)

XCEL_sheet_set_name($workbook; 1; "New Sheet")
$sheet:=XCEL_sheet_get_name($workbook; 1)

$sheet:=XCEL_sheet_create($workbook)
XCEL_sheet_set_name($workbook; 2; "Even New Sheet")

$image_path_t:=Get 4D folder:C485(Current resources folder:K5:16)+"Bluehound.gif"
XCEL_sheet_set_background($workbook; $sheet; $image_path_t)
XCEL_workbook_close($workbook)

//tested on Mac: OK
//tested on Win: OK

TRACE:C157