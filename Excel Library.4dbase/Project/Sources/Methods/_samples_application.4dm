//%attributes = {"invisible":true}
//methods:
//http://msdn.microsoft.com/en-us/library/bb149134(v=office.12).aspx
//properties:
//http://msdn.microsoft.com/en-us/library/dd787731(v=office.12).aspx

XCEL_workbook_close_all
$workbook:=XCEL_workbook_create
XCEL_application_show

$count_l:=XCEL_sheet_count($workbook)

XCEL_sheet_set_name($workbook; 1; "Application Example")
$worksheet:=XCEL_sheet_get_name($workbook; 1)

//DisplayFullScreen
//http://msdn.microsoft.com/en-us/library/bb177506(v=office.12).aspx
XCEL_application_set_fullscreen(True:C214)
$fullscreen_b:=XCEL_application_get_fullscreen
XCEL_range_set_value($workbook; $worksheet; "A1"; "DisplayFullScreen")
XCEL_range_set_value($workbook; $worksheet; "B1"; String:C10($fullscreen_b))

//Left
//http://msdn.microsoft.com/en-us/library/bb179255(v=office.12).aspx
//Top
//http://msdn.microsoft.com/en-us/library/bb214199(v=office.12).aspx
//Width
//http://msdn.microsoft.com/en-us/library/bb214430(v=office.12).aspx
//Height
//http://msdn.microsoft.com/en-us/library/bb179253(v=office.12).aspx
C_REAL:C285($left_r; $top_r; $width_r; $height_r)
XCEL_application_set_rect(10; 10; 300; 400)
XCEL_application_get_rect(->$left_r; ->$top_r; ->$width_r; ->$height_r)
XCEL_range_set_value($workbook; $worksheet; "A2"; "Left")
XCEL_range_set_value($workbook; $worksheet; "B2"; String:C10($left_r))
XCEL_range_set_value($workbook; $worksheet; "A3"; "Top")
XCEL_range_set_value($workbook; $worksheet; "B3"; String:C10($top_r))
XCEL_range_set_value($workbook; $worksheet; "A4"; "Right")
XCEL_range_set_value($workbook; $worksheet; "B4"; String:C10($width_r))
XCEL_range_set_value($workbook; $worksheet; "A5"; "Height")
XCEL_range_set_value($workbook; $worksheet; "B5"; String:C10($height_r))

//note:  the main application on Windows, the active window on Mac

TRACE:C157

XCEL_workbook_close($workbook)
XCEL_application_hide

//test on both Mac and PC