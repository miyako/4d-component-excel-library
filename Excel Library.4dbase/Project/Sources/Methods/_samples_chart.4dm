//%attributes = {"invisible":true}
//methods:
//http://msdn.microsoft.com/en-us/library/bb244251(v=office.12).aspx
//properties:
//http://msdn.microsoft.com/en-us/library/bb244250(v=office.12).aspx

XCEL_workbook_CLOSE_ALL
$workbook:=XCEL_workbook_Create
XCEL_application_SHOW

$count_l:=XCEL_sheet_Count($workbook)

XCEL_sheet_SET_NAME($workbook; 1; "Chart Example")
$worksheet:=XCEL_sheet_Get_name($workbook; 1)

XCEL_range_SET_VALUE($workbook; $worksheet; "A1:A5"; "10")
$chart:=XCEL_chart_Create($workbook; $worksheet; "A1:A5"; "My Chart"; 10; 200; 400; 250)
$count_l:=XCEL_chart_Count($workbook; $worksheet)

//Name
//http://msdn.microsoft.com/en-us/library/bb179461(v=office.12).aspx
XCEL_chart_SET_NAME($workbook; $worksheet; 1; "new chart")
$chart:=XCEL_chart_Get_name($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A7"; "Name")
XCEL_range_SET_VALUE($workbook; $worksheet; "B7"; $chart)

//ChartType
//http://msdn.microsoft.com/en-us/library/bb179424(v=office.12).aspx
XCEL_chart_SET_TYPE($workbook; $worksheet; $chart; -4120)
$type_l:=XCEL_chart_Get_type($workbook; $worksheet; $chart)
XCEL_range_SET_VALUE($workbook; $worksheet; "A8"; "Type")
XCEL_range_SET_VALUE($workbook; $worksheet; "B8"; String:C10($type_l))

//note: see MSDN documentation for full list of chart types 
//http://msdn.microsoft.com/en-us/library/bb241008.aspx

//picture appearance enumerartion
//http://msdn.microsoft.com/en-us/library/bb241413(v=office.12).aspx
//copy picture format enumerartion
//http://msdn.microsoft.com/en-us/library/bb241043(v=office.12).aspx
XCEL_chart_COPY_PICTURE($workbook; $worksheet; $chart; XCEL_picture_appearance_Printer; XCEL_picture_format_Picture)
//the chart image is in the pasteboard

TRACE:C157

XCEL_workbook_CLOSE($workbook)
XCEL_application_HIDE

//test on both Mac and PC