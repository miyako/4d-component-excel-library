//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($2)
C_TEXT:C284($3)
C_LONGINT:C283($4)
C_POINTER:C301(${5})

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_SHEET_NAME"; $2)
SET ENVIRONMENT VARIABLE:C812("XCEL_RANGE"; $3)
SET ENVIRONMENT VARIABLE:C812("XCEL_RANGE_BORDER_TYPE"; String:C10($4))

$json_t:=XCEL_util_EXECUTE("range_get_border")

ARRAY LONGINT:C221($match_position_al; 0)
ARRAY LONGINT:C221($match_length_al; 0)

If (Match regex:C1019("\\{\"style\":(.+?),\"weight\":(.+?),\"color\":(.+?)\\}"; $json_t; 1; $match_position_al; $match_length_al))
	
	C_LONGINT:C283($style_l; $weight_l; $color_l)
	
	$style_l:=Num:C11(Substring:C12($json_t; $match_position_al{1}; $match_length_al{1}))
	$weight_l:=Num:C11(Substring:C12($json_t; $match_position_al{2}; $match_length_al{2}))
	$color_l:=Num:C11(Substring:C12($json_t; $match_position_al{3}; $match_length_al{3}))
	
	XCEL_util_GET_PARAMETER($5; ->$style_l)
	XCEL_util_GET_PARAMETER($6; ->$weight_l)
	XCEL_util_GET_PARAMETER($7; ->$color_l)
	
End if 