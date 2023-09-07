//%attributes = {"invisible":true}
//methods:
//http://msdn.microsoft.com/en-us/library/bb259321(v=office.12).aspx
//properties:
//http://msdn.microsoft.com/en-us/library/bb259311(v=office.12).aspx

$image_path_t:=Get 4D folder:C485(Current resources folder:K5:16)+"4D.png"

XCEL_workbook_CLOSE_ALL
$workbook:=XCEL_workbook_Create
XCEL_application_SHOW

XCEL_sheet_SET_NAME($workbook; 1; "shape demo")
$sheet:=XCEL_sheet_Get_name($workbook; 1)

//the 'shape' class represent any shape object in Excel
//note that 'picture' is a sub-class of shape;
//all shape functions work with pictures too.

$picture:=XCEL_picture_Create($workbook; $sheet; $image_path_t; 100; 100; 128; 128)
$count_l:=XCEL_shape_Count($workbook; $sheet)
XCEL_shape_SET_NAME($workbook; $sheet; 1; "new picture")
$picture:=XCEL_shape_Get_name($workbook; $sheet; 1)

XCEL_shape_SET_ROTATION($workbook; $sheet; $picture; 45)
$rotation_r:=XCEL_shape_Get_rotation($workbook; $sheet; $picture)

//chart picture placement enumeration
//http://msdn.microsoft.com/en-us/library/bb241002(v=office.12).aspx
XCEL_shape_SET_PLACEMENT($workbook; $sheet; $picture; XCEL_placement_Move)
$placement_l:=XCEL_shape_Get_placement($workbook; $sheet; $picture)

XCEL_shape_SET_LOCK_RATIO($workbook; $sheet; $picture; True:C214)
$lock_aspect_ratio_b:=XCEL_shape_Get_lock_ratio($workbook; $sheet; $picture)

C_REAL:C285($left_r; $top_r; $width_r; $height_r)
XCEL_shape_SET_RECT($workbook; $sheet; $picture; 10; 10; 64; 64)
XCEL_shape_GET_RECT($workbook; $sheet; $picture; ->$left_r; ->$top_r; ->$width_r; ->$height_r)

XCEL_workbook_CLOSE($workbook)

//tested on Mac: OK
//tested on Win: OK

TRACE:C157