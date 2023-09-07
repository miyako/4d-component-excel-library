//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($2)
C_TEXT:C284($3)
C_POINTER:C301(${4})

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_SHEET_NAME"; $2)
SET ENVIRONMENT VARIABLE:C812("XCEL_RANGE"; $3)

$json_t:=XCEL_util_EXECUTE("range_get_rect")

ARRAY LONGINT:C221($match_position_al; 0)
ARRAY LONGINT:C221($match_length_al; 0)

If (Match regex:C1019("\\{\"left\":(.+?),\"top\":(.+?),\"width\":(.+?),\"height\":(.+?)\\}"; $json_t; 1; $match_position_al; $match_length_al))
	
	C_REAL:C285($left_r; $top_r; $width_r; $height_r)
	
	$left_r:=Num:C11(Substring:C12($json_t; $match_position_al{1}; $match_length_al{1}))
	$top_r:=Num:C11(Substring:C12($json_t; $match_position_al{2}; $match_length_al{2}))
	$width_r:=Num:C11(Substring:C12($json_t; $match_position_al{3}; $match_length_al{3}))
	$height_r:=Num:C11(Substring:C12($json_t; $match_position_al{4}; $match_length_al{4}))
	
	XCEL_util_GET_PARAMETER($4; ->$left_r)
	XCEL_util_GET_PARAMETER($5; ->$top_r)
	XCEL_util_GET_PARAMETER($6; ->$width_r)
	XCEL_util_GET_PARAMETER($7; ->$height_r)
	
End if 