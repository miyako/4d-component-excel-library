//%attributes = {"invisible":true}
//properties:
//http://msdn.microsoft.com/en-us/library/bb258982(v=office.12).aspx

XCEL_workbook_CLOSE_ALL
$workbook:=XCEL_workbook_Create
XCEL_application_SHOW

$count_l:=XCEL_sheet_Count($workbook)

XCEL_sheet_SET_NAME($workbook; 1; "Page Setup Example")
$worksheet:=XCEL_sheet_Get_name($workbook; 1)

//you can get/set the following page setup properties.

//note that page setup is a property of a worksheet;
//therefore you need to specify the workbook name and worksheet name,
//to call any of these functions.

//FitToPagesTall
//http://msdn.microsoft.com/en-us/library/bb208514(v=office.12).aspx
XCEL_setup_SET_FIT_PAGES_TALL($workbook; $worksheet; 2)
$fit_to_pages_tall_l:=XCEL_setup_Get_fit_pages_tall($workbook; $worksheet)
XCEL_range_SET_VALUE($workbook; $worksheet; "A1"; "FitToPagesTall")
XCEL_range_SET_VALUE($workbook; $worksheet; "B1"; String:C10($fit_to_pages_tall_l))

//FitToPagesWide
//http://msdn.microsoft.com/en-us/library/bb208515(v=office.12).aspx
XCEL_setup_SET_FIT_PAGES_WIDE($workbook; $worksheet; 3)
$fit_to_pages_wide_l:=XCEL_setup_Get_fit_pages_wide($workbook; $worksheet)
XCEL_range_SET_VALUE($workbook; $worksheet; "A2"; "FitToPagesWide")
XCEL_range_SET_VALUE($workbook; $worksheet; "B2"; String:C10($fit_to_pages_wide_l))

//note: 'FitToPages___' on Windows, the propery is variant, in that it can be boolean as well as integer.
//for cross platform compatibility we will only support interger.

//Zoom
//http://msdn.microsoft.com/en-us/library/bb214929(v=office.12).aspx
XCEL_setup_SET_ZOOM($workbook; $worksheet; 200)
$zoom_r:=XCEL_setup_Get_zoom($workbook; $worksheet)
XCEL_range_SET_VALUE($workbook; $worksheet; "A3"; "Zoom")
XCEL_range_SET_VALUE($workbook; $worksheet; "B3"; String:C10($zoom_r))

//note: 'Zoom' is variant on both platforms, it can be boolean as well as numeric percentages between 10 and 400.
//for simplicity we will only support real.

//FirstPageNumber
//http://msdn.microsoft.com/en-us/library/bb208512(v=office.12).aspx
XCEL_setup_SET_1ST_PAGE_NUMBER($workbook; $worksheet; 4)
$first_page_number_l:=XCEL_setup_Get_1st_page_number($workbook; $worksheet)
XCEL_range_SET_VALUE($workbook; $worksheet; "A4"; "FirstPageNumber")
XCEL_range_SET_VALUE($workbook; $worksheet; "B4"; String:C10($first_page_number_l))

//Orientation
//http://msdn.microsoft.com/en-us/library/bb213219(v=office.12).aspx
//constant xlLandscape (2)
XCEL_setup_SET_ORIENTATION($workbook; $worksheet; XCEL_page_orientation_Landscape)
$orientation_l:=XCEL_setup_Get_orientation($workbook; $worksheet)
XCEL_range_SET_VALUE($workbook; $worksheet; "A5"; "Orientation")
XCEL_range_SET_VALUE($workbook; $worksheet; "B5"; String:C10($orientation_l))
//constant xlPortrait (1)
XCEL_setup_SET_ORIENTATION($workbook; $worksheet; XCEL_page_orientation_Portrait)
$orientation_l:=XCEL_setup_Get_orientation($workbook; $worksheet)
XCEL_range_SET_VALUE($workbook; $worksheet; "A6"; "Orientation")
XCEL_range_SET_VALUE($workbook; $worksheet; "B6"; String:C10($orientation_l))

//TopMargin
//http://msdn.microsoft.com/en-us/library/bb221846(v=office.12).aspx
XCEL_setup_SET_TOP_MARGIN_CM($workbook; $worksheet; 2.54)
$top_margin_r:=XCEL_setup_Get_top_margin($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A7"; "TopMargin")
XCEL_range_SET_VALUE($workbook; $worksheet; "B7"; String:C10($top_margin_r))
XCEL_setup_SET_TOP_MARGIN_IN($workbook; $worksheet; 1)
$top_margin_r:=XCEL_setup_Get_top_margin($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A8"; "TopMargin")
XCEL_range_SET_VALUE($workbook; $worksheet; "B8"; String:C10($top_margin_r))

//note: margins can be set by centimeters or inches, but be aware that the internal format is always points
//1 inch is 72 points on a 72 dpi environment
//1 inch is 2.54 centimeters, hence 1 centimeter is (2.54*72) points.
//however there ican be a margin of error (no pun intended) if you set decimal centimeters

//LeftMargin
//http://msdn.microsoft.com/en-us/library/bb177853(v=office.12).aspx
XCEL_setup_SET_LEFT_MARGIN_CM($workbook; $worksheet; 2.54)
$left_margin_r:=XCEL_setup_Get_left_margin($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A9"; "LeftMargin")
XCEL_range_SET_VALUE($workbook; $worksheet; "B9"; String:C10($left_margin_r))
XCEL_setup_SET_LEFT_MARGIN_IN($workbook; $worksheet; 1)
$left_margin_r:=XCEL_setup_Get_left_margin($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A10"; "LeftMargin")
XCEL_range_SET_VALUE($workbook; $worksheet; "B10"; String:C10($left_margin_r))

//RightMargin
//http://msdn.microsoft.com/en-us/library/bb209175(v=office.12).aspx
XCEL_setup_SET_RIGHT_MARGIN_CM($workbook; $worksheet; 2.54)
$right_margin_r:=XCEL_setup_Get_right_margin($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A11"; "RightMargin")
XCEL_range_SET_VALUE($workbook; $worksheet; "B11"; String:C10($right_margin_r))
XCEL_setup_SET_RIGHT_MARGIN_IN($workbook; $worksheet; 1)
$right_margin_r:=XCEL_setup_Get_right_margin($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A12"; "RightMargin")
XCEL_range_SET_VALUE($workbook; $worksheet; "B12"; String:C10($right_margin_r))

//BottomMargin
//http://msdn.microsoft.com/en-us/library/bb220892(v=office.12).aspx
XCEL_setup_SET_BOTTOM_MARGIN_CM($workbook; $worksheet; 2.54)
$bottom_margin_r:=XCEL_setup_Get_bottom_margin($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A13"; "BottomMargin")
XCEL_range_SET_VALUE($workbook; $worksheet; "B13"; String:C10($bottom_margin_r))
XCEL_setup_SET_BOTTOM_MARGIN_IN($workbook; $worksheet; 1)
XCEL_setup_SET_BOTTOM_MARGIN_CM($workbook; $worksheet; 2.54)
XCEL_range_SET_VALUE($workbook; $worksheet; "A14"; "BottomMargin")
XCEL_range_SET_VALUE($workbook; $worksheet; "B14"; String:C10($bottom_margin_r))

//HeaderMargin
//http://msdn.microsoft.com/en-us/library/bb208664(v=office.12).aspx
XCEL_setup_SET_HEADER_MARGIN_CM($workbook; $worksheet; 2.54)
$header_margin_r:=XCEL_setup_Get_header_margin($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A15"; "HeaderMargin")
XCEL_range_SET_VALUE($workbook; $worksheet; "B15"; String:C10($header_margin_r))
XCEL_setup_SET_HEADER_MARGIN_IN($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A16"; "HeaderMargin")
XCEL_range_SET_VALUE($workbook; $worksheet; "B16"; String:C10($header_margin_r))

//FooterMargin
//http://msdn.microsoft.com/en-us/library/bb208526(v=office.12).aspx
XCEL_setup_SET_FOOTER_MARGIN_CM($workbook; $worksheet; 2.54)
$footer_margin_r:=XCEL_setup_Get_footer_margin($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A17"; "FooterMargin")
XCEL_range_SET_VALUE($workbook; $worksheet; "B17"; String:C10($footer_margin_r))
XCEL_setup_SET_FOOTER_MARGIN_IN($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A18"; "FooterMargin")
XCEL_range_SET_VALUE($workbook; $worksheet; "B18"; String:C10($footer_margin_r))

//BlackAndWhite
//http://msdn.microsoft.com/en-us/library/bb220890(v=office.12).aspx
XCEL_setup_SET_BLACK_AND_WHITE($workbook; $worksheet; True:C214)
$black_and_white_b:=XCEL_setup_Get_black_and_white($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A19"; "BlackAndWhite")
XCEL_range_SET_VALUE($workbook; $worksheet; "B19"; String:C10(Num:C11($black_and_white_b); "True;;False"))
XCEL_setup_SET_BLACK_AND_WHITE($workbook; $worksheet; False:C215)
$black_and_white_b:=XCEL_setup_Get_black_and_white($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A20"; "BlackAndWhite")
XCEL_range_SET_VALUE($workbook; $worksheet; "B20"; String:C10(Num:C11($black_and_white_b); "True;;False"))

//note: although defined as bool, the 'BlackAndWhite' property must be explicitly cast on Windows

//Draft
//http://msdn.microsoft.com/en-us/library/bb221052(v=office.12).aspx
XCEL_setup_SET_DRAFT($workbook; $worksheet; True:C214)
$draft_b:=XCEL_setup_Get_draft($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A21"; "Draft")
XCEL_range_SET_VALUE($workbook; $worksheet; "B21"; String:C10(Num:C11($draft_b); "True;;False"))
XCEL_setup_SET_DRAFT($workbook; $worksheet; False:C215)
$draft_b:=XCEL_setup_Get_draft($workbook; $worksheet; 1)
XCEL_range_SET_VALUE($workbook; $worksheet; "A22"; "Draft")
XCEL_range_SET_VALUE($workbook; $worksheet; "B22"; String:C10(Num:C11($draft_b); "True;;False"))

TRACE:C157

XCEL_workbook_CLOSE($workbook)

//test on both Mac and PC