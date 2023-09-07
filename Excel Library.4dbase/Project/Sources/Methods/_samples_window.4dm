//%attributes = {"invisible":true}
//methods:
//http://msdn.microsoft.com/en-us/library/bb245598(v=office.12).aspx
//properties:
//http://msdn.microsoft.com/en-us/library/bb245593(v=office.12).aspx

XCEL_workbook_CLOSE_ALL
$workbook:=XCEL_workbook_Create
XCEL_application_SHOW

$count_l:=XCEL_window_Count($workbook)
XCEL_window_SET_CAPTION($workbook; 1; "new window")
$window:=XCEL_window_Get_caption($workbook; 1)

XCEL_window_SET_VISIBLE($workbook; $window; False:C215)
$visible_b:=XCEL_window_Get_visible($workbook; $window)

XCEL_window_SET_ZOOM($workbook; $window; 200)
$zoom_b:=XCEL_window_Get_zoom($workbook; $window)

XCEL_window_SET_VIEW($workbook; $window; XCEL_View_normal)
XCEL_window_SET_VIEW($workbook; $window; XCEL_View_page_layout)
$view_b:=XCEL_window_Get_view($workbook; $window)

XCEL_window_SET_ENABLE_RESIZE($workbook; $window; True:C214)
$enable_resize_b:=XCEL_window_Get_enable_resize($workbook; $window)

XCEL_window_SET_FREEZE_PANES($workbook; $window; True:C214)
$freeze_panes_b:=XCEL_window_Get_freeze_panes($workbook; $window)

XCEL_window_SET_GRIDLINE_COLOR($workbook; $window; 5)
$gridline_color_l:=XCEL_window_Get_gridline_color($workbook; $window)

C_REAL:C285($left_r; $top_r; $width_r; $height_r)
XCEL_window_SET_RECT($workbook; $window; 0; 0; 100; 100)
XCEL_window_GET_RECT($workbook; $window; ->$left_r; ->$top_r; ->$width_r; ->$height_r)

XCEL_window_SET_SHOW_GRIDLINES($workbook; $window; True:C214)
$gridlines_b:=XCEL_window_Get_show_gridlines($workbook; $window)

XCEL_window_SET_SHOW_HEADINGS($workbook; $window; True:C214)
$headings_b:=XCEL_window_Get_show_headings($workbook; $window)

XCEL_window_SET_SHOW_OUTLINE($workbook; $window; True:C214)
$outline_b:=XCEL_window_Get_show_outline($workbook; $window)

XCEL_window_SET_SHOW_SCROLL_H($workbook; $window; True:C214)
$scroll_h_b:=XCEL_window_Get_show_scroll_h($workbook; $window)

XCEL_window_SET_SHOW_SCROLL_V($workbook; $window; True:C214)
$scroll_v_b:=XCEL_window_Get_show_scroll_v($workbook; $window)

XCEL_window_SET_SHOW_TABS($workbook; $window; True:C214)
$tabs_b:=XCEL_window_Get_show_tabs($workbook; $window)

XCEL_window_SET_SPLIT_COLUMN($workbook; $window; 10)
$split_column_b:=XCEL_window_Get_split_column($workbook; $window)

XCEL_window_SET_SPLIT_ROW($workbook; $window; 10)
$row_b:=XCEL_window_Get_split_row($workbook; $window)

XCEL_window_SET_TAB_RATIO($workbook; $window; 10)
$tab_ratio_b:=XCEL_window_Get_tab_ratio($workbook; $window)

XCEL_application_ARRANGE_WINDOW(XCEL_arrange_style_Cascade)
XCEL_application_ARRANGE_WINDOW(XCEL_arrange_style_Horizontal)
XCEL_application_ARRANGE_WINDOW(XCEL_arrange_style_Tiled)
XCEL_application_ARRANGE_WINDOW(XCEL_arrange_style_Vertical)
