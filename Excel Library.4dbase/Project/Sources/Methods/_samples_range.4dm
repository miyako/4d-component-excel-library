//%attributes = {"invisible":true}
//methods:
//http://msdn.microsoft.com/en-us/library/bb245315(v=office.12).aspx
//properties:
//http://msdn.microsoft.com/en-us/library/bb245304(v=office.12).aspx

XCEL_workbook_CLOSE_ALL
$workbook:=XCEL_workbook_Create
XCEL_application_SHOW

$count_l:=XCEL_sheet_Count($workbook)

XCEL_sheet_SET_NAME($workbook; 1; "Range Example")
$worksheet:=XCEL_sheet_Get_name($workbook; 1)

//Insert
//http://www.microsoft.com/japan/technet/scriptcenter/resources/qanda/apr05/hey0411.mspx
XCEL_range_SET_VALUE($workbook; $worksheet; "a1"; "insert")
XCEL_range_INSERT($workbook; $worksheet; "a1"; XCEL_shift_Right)
XCEL_range_INSERT($workbook; $worksheet; "b1"; XCEL_shift_Down)

//Delete
//http://msdn.microsoft.com/en-us/library/bb178843(v=office.12).aspx
XCEL_range_DELETE($workbook; $worksheet; "b1"; XCEL_shift_Up)
XCEL_range_DELETE($workbook; $worksheet; "a1"; XCEL_shift_Left)

//FillUp, FillDown, FillLeft, FillRight 
//http://msdn.microsoft.com/en-us/library/bb209849(v=office.12).aspx
XCEL_range_SET_VALUE($workbook; $worksheet; "a1"; "fill down")
XCEL_range_FILL_DOWN($workbook; $worksheet; "a1:a10")
XCEL_range_SET_VALUE($workbook; $worksheet; "b10"; "fill up")
XCEL_range_FILL_UP($workbook; $worksheet; "b10:b1")
XCEL_range_SET_VALUE($workbook; $worksheet; "c3"; "fill right")
XCEL_range_FILL_RIGHT($workbook; $worksheet; "c3:g1")
XCEL_range_SET_VALUE($workbook; $worksheet; "g2"; "fill left")
XCEL_range_FILL_LEFT($workbook; $worksheet; "g2:c2")

//auto fill type Enumeration
//http://msdn.microsoft.com/en-us/library/bb240952(v=office.12).aspx

//AutoFill
//http://msdn.microsoft.com/en-us/library/bb209671(v=office.12).aspx

XCEL_range_SET_VALUE($workbook; $worksheet; "a11"; "monday")
XCEL_range_AUTOFILL($workbook; $worksheet; "a11"; "a11:a17"; XCEL_fill_Weekdays)
XCEL_range_SET_VALUE($workbook; $worksheet; "b11"; "january")
XCEL_range_AUTOFILL($workbook; $worksheet; "b11"; "b11:b22"; XCEL_fill_Months)

//NumberFormat
//http://msdn.microsoft.com/en-us/library/bb213677(v=office.12).aspx
XCEL_range_SET_NUMBER_FORMAT($workbook; $worksheet; "A1"; "$#,##0.00_);[Red]($#,##0.00)")
XCEL_range_SET_NUMBER_FORMAT($workbook; $worksheet; "A2"; "hh:mm:ss")
$number_format_t:=XCEL_range_Get_number_format($workbook; $worksheet; "A1")
$number_format_t:=XCEL_range_Get_number_format($workbook; $worksheet; "A2")

//Borders
//http://msdn.microsoft.com/en-us/library/bb213512(v=office.12).aspx
XCEL_range_SET_BORDER($workbook; $worksheet; "B1:B9"; XCEL_border_Bottom; XCEL_border_style_Dash; XCEL_border_weight_Medium; 7)
XCEL_range_SET_BORDER($workbook; $worksheet; "B1:B9"; XCEL_border_Left; XCEL_border_style_Dash; XCEL_border_weight_Medium; 7)
XCEL_range_SET_BORDER($workbook; $worksheet; "B1:B9"; XCEL_border_Right; XCEL_border_style_Dash; XCEL_border_weight_Medium; 7)
XCEL_range_SET_BORDER($workbook; $worksheet; "B1:B9"; XCEL_border_Top; XCEL_border_style_Dash; XCEL_border_weight_Medium; 7)

C_REAL:C285($left_r; $top_r; $width_r; $height_r)
XCEL_range_GET_RECT($workbook; $worksheet; "B1:B9"; ->$left_r; ->$top_r; ->$width_r; ->$height_r)

C_LONGINT:C283($style_l; $weight_l; $color_l)

XCEL_range_GET_BORDER($workbook; $worksheet; "B1"; XCEL_border_Bottom; ->$style_l; ->$weight_l; ->$color_l)

ARRAY TEXT:C222($values_at; 5)
$values_at{1}:="1"
$values_at{2}:="2"
$values_at{3}:="3"
$values_at{4}:="4"
$values_at{5}:="5"

//running individual scripts for each cell can be expensive;
XCEL_range_SET_VALUE_ARRAY($workbook; $worksheet; "A1:A7"; ->$values_at)

//Value
//http://msdn.microsoft.com/en-us/library/bb238606(v=office.12).aspx
XCEL_range_SET_VALUE($workbook; $worksheet; "a1"; "bold")

XCEL_range_SET_FONT_BOLD($workbook; $worksheet; "a1"; True:C214)
$bold_b:=XCEL_range_Get_font_bold($workbook; $worksheet; "a1")

XCEL_range_SET_VALUE($workbook; $worksheet; "a2"; "italic")
XCEL_range_SET_FONT_ITALIC($workbook; $worksheet; "a2"; True:C214)
$italic_b:=XCEL_range_Get_font_italic($workbook; $worksheet; "a2")

XCEL_range_SET_VALUE($workbook; $worksheet; "a3"; "color #2")
XCEL_range_SET_FONT_COLOR($workbook; $worksheet; "a3"; 3)
$color_l:=XCEL_range_Get_font_color($workbook; $worksheet; "a3")

XCEL_range_SET_VALUE($workbook; $worksheet; "a4"; "Courier")
XCEL_range_SET_FONT_NAME($workbook; $worksheet; "a4"; "Courier")
$font_name_t:=XCEL_range_Get_font_name($workbook; $worksheet; "a4")

XCEL_range_SET_VALUE($workbook; $worksheet; "a5"; "outline")
XCEL_range_SET_FONT_OUTLINE($workbook; $worksheet; "a5"; True:C214)
$outline_b:=XCEL_range_Get_font_outline($workbook; $worksheet; "a5")

XCEL_range_SET_VALUE($workbook; $worksheet; "a6"; "shadow")
XCEL_range_SET_FONT_SHADOW($workbook; $worksheet; "a6"; True:C214)
$shadow_b:=XCEL_range_Get_font_shadow($workbook; $worksheet; "a6")

XCEL_range_SET_VALUE($workbook; $worksheet; "a7"; "size 20")
XCEL_range_SET_FONT_SIZE($workbook; $worksheet; "a7"; 20)
$size_r:=XCEL_range_Get_font_size($workbook; $worksheet; "a7")

XCEL_range_SET_VALUE($workbook; $worksheet; "a8"; "strike through")
XCEL_range_SET_STRIKE_THROUGH($workbook; $worksheet; "a8"; True:C214)
$strike_through_b:=XCEL_range_Get_strike_through($workbook; $worksheet; "a8")

XCEL_range_SET_VALUE($workbook; $worksheet; "a9"; "subscript")
XCEL_range_SET_SUBSCRIPT($workbook; $worksheet; "a9"; True:C214)
$subscript_b:=XCEL_range_Get_subscript($workbook; $worksheet; "a9")

XCEL_range_SET_VALUE($workbook; $worksheet; "a10"; "superscript")
XCEL_range_SET_SUPERSCRIPT($workbook; $worksheet; "a10"; True:C214)
$superscript_b:=XCEL_range_Get_superscript($workbook; $worksheet; "a10")

//actually, 'value' is interpreted as formula, but it is better to set explicitly
XCEL_range_SET_FORMULA($workbook; $worksheet; "a11"; "=TODAY()")
$formula_t:=XCEL_range_Get_formula($workbook; $worksheet; "a11")

XCEL_range_SET_VALUE($workbook; $worksheet; "a12"; "height 15")
XCEL_range_SET_HEIGHT($workbook; $worksheet; "a12"; 15)
$height_r:=XCEL_range_Get_height($workbook; $worksheet; "a12")

XCEL_range_SET_VALUE($workbook; $worksheet; "a13"; "interior color #5")
XCEL_range_SET_INTERIOR_COLOR($workbook; $worksheet; "a13"; 5)
$interior_color_l:=XCEL_range_Get_interior_color($workbook; $worksheet; "a13")

XCEL_range_SET_VALUE($workbook; $worksheet; "a14"; "locked")
XCEL_range_SET_LOCKED($workbook; $worksheet; "a14"; True:C214)
$locked_b:=XCEL_range_Get_locked($workbook; $worksheet; "a14")

//there are naming rules, suspect the name if command fails
XCEL_range_SET_VALUE($workbook; $worksheet; "a15"; "4")
XCEL_range_SET_NAME($workbook; $worksheet; "a15"; "name")
$name_t:=XCEL_range_Get_name($workbook; $worksheet; "a15")

XCEL_range_SET_VALUE($workbook; $worksheet; "a16"; "orientation 45")
XCEL_range_SET_ORIENTATION($workbook; $worksheet; "a16"; 45)
$orientation_r:=XCEL_range_Get_orientation($workbook; $worksheet; "a16")

XCEL_range_SET_VALUE($workbook; $worksheet; "a19"; "width 20")
XCEL_range_SET_WIDTH($workbook; $worksheet; "a19"; 15)
$width_r:=XCEL_range_Get_width($workbook; $worksheet; "a19")

XCEL_range_SET_VALUE($workbook; $worksheet; "a20"; "wrap text")
XCEL_range_SET_WRAP_TEXT($workbook; $worksheet; "a20"; True:C214)
$wrap_text_b:=XCEL_range_Get_wrap_text($workbook; $worksheet; "a20")

XCEL_range_SET_VALUE($workbook; $worksheet; "a22"; "shrink to fit")
XCEL_range_SET_SHRINK_TO_FIT($workbook; $worksheet; "a22"; True:C214)
$shrink_to_fit_b:=XCEL_range_Get_shrink_to_fit($workbook; $worksheet; "a22")

//underline style enumeration
//http://msdn.microsoft.com/en-us/library/bb216406(v=office.12).aspx
XCEL_range_SET_VALUE($workbook; $worksheet; "a23"; "double")
XCEL_range_SET_UNDERLINE($workbook; $worksheet; "a23"; XCEL_underline_Double)
XCEL_range_SET_VALUE($workbook; $worksheet; "a24"; "double accounting")
XCEL_range_SET_UNDERLINE($workbook; $worksheet; "a24"; XCEL_underline_Double_account)
XCEL_range_SET_VALUE($workbook; $worksheet; "a25"; "single")
XCEL_range_SET_UNDERLINE($workbook; $worksheet; "a25"; XCEL_underline_Single)
XCEL_range_SET_VALUE($workbook; $worksheet; "a26"; "single accounting")
XCEL_range_SET_UNDERLINE($workbook; $worksheet; "a26"; XCEL_underline_Single_account)
XCEL_range_SET_VALUE($workbook; $worksheet; "a27"; "none")
XCEL_range_SET_UNDERLINE($workbook; $worksheet; "a27"; XCEL_underline_None)
$underline_l:=XCEL_range_Get_underline($workbook; $worksheet; "a27")

XCEL_range_SET_VALUE($workbook; $worksheet; "a28"; "merge")
XCEL_range_MERGE($workbook; $worksheet; "a28:a30")
XCEL_range_UNMERGE($workbook; $worksheet; "a28:a30")
XCEL_range_MERGE_ACROSS($workbook; $worksheet; "a28:c28")

//picture appearance enumerartion
//http://msdn.microsoft.com/en-us/library/bb241413(v=office.12).aspx
//copy picture format enumerartion
//http://msdn.microsoft.com/en-us/library/bb241043(v=office.12).aspx
XCEL_range_COPY_PICTURE($workbook; $worksheet; "a1:a5"; XCEL_picture_appearance_Printer; XCEL_picture_format_Picture)
XCEL_workbook_CLOSE($workbook)

//tested on Mac: OK
//tested on Win: OK

TRACE:C157