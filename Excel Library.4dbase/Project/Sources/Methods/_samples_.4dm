//%attributes = {"invisible":true}
XCEL_workbook_CLOSE_ALL
$workbook:=XCEL_workbook_Create
XCEL_application_SHOW
$worksheet:=XCEL_sheet_Get_name($workbook; 1)

ARRAY TEXT:C222($values_at; 20)
$values_at{1}:="automatic"
$values_at{2}:="checker"
$values_at{3}:="criss cross"
$values_at{4}:="down"
$values_at{5}:="gray 16"

$values_at{6}:="gray 25"
$values_at{7}:="gray 50"
$values_at{8}:="gray 75"
$values_at{9}:="gray 8"
$values_at{10}:="grid"

$values_at{11}:="horizontal"
$values_at{12}:="light down"
$values_at{13}:="light horizontal"
$values_at{14}:="light up"
$values_at{15}:="light vertical"

$values_at{16}:="none"
$values_at{17}:="semi gray 75"
$values_at{18}:="solid"
$values_at{19}:="up"
$values_at{20}:="vertical"

XCEL_range_SET_VALUE_ARRAY($workbook; $worksheet; "A1:A20"; ->$values_at)

XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A1"; XCEL_pattern_Automatic; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A2"; XCEL_pattern_Checker; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A3"; XCEL_pattern_Criss_cross; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A4"; XCEL_pattern_Down; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A5"; XCEL_pattern_Gray_16; 4)

XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A6"; XCEL_pattern_Gray_25; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A7"; XCEL_pattern_Gray_50; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A8"; XCEL_pattern_Gray_75; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A9"; XCEL_pattern_Gray_8; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A10"; XCEL_pattern_Grid; 4)

XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A11"; XCEL_pattern_Horizontal; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A12"; XCEL_pattern_Light_down; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A13"; XCEL_pattern_Light_horizontal; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A14"; XCEL_pattern_Light_up; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A15"; XCEL_pattern_Light_vertical; 4)

XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A16"; XCEL_pattern_None; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A17"; XCEL_pattern_Semi_gray_75; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A18"; XCEL_pattern_Solid; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A19"; XCEL_pattern_Up; 4)
XCEL_range_SET_INTERIOR_PATTERN($workbook; $worksheet; "A20"; XCEL_pattern_Vertical; 4)

C_LONGINT:C283($XCEL_pattern_l; $color_l)

XCEL_range_GET_INTERIOR_PATTERN($workbook; $worksheet; "A5"; ->$XCEL_pattern_l; ->$color_l)

TRACE:C157

XCEL_workbook_CLOSE($workbook)
XCEL_application_HIDE