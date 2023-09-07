//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($2)
C_TEXT:C284($3)

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_MAIL_RECIPIENTS"; $2)  //semicolon delimiter
SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_MAIL_SUBJECT"; $3)

C_LONGINT:C283($platform_l)
_O_PLATFORM PROPERTIES:C365($platform_l)

If ($platform_l=Windows:K25:3)
	SET ENVIRONMENT VARIABLE:C812("_4D_OPTION_BLOCKING_EXTERNAL_PROCESS"; "False")
Else 
	
	C_LONGINT:C283($i; $recipients_l)
	
	$i:=1
	$recipients_l:=0
	
	ARRAY LONGINT:C221($match_position_al; 0)
	ARRAY LONGINT:C221($match_length_al; 0)
	
	While (Match regex:C1019("(.+?)(;|$)"; $2; $i; $match_position_al; $match_length_al))
		$recipients_l:=$recipients_l+1
		SET ENVIRONMENT VARIABLE:C812("XCEL_MAIL_RECIPIENT"+String:C10($recipients_l); Substring:C12($2; $match_position_al{1}; $match_length_al{1}))
		$i:=$match_position_al{1}+$match_length_al{1}
	End while 
	
	SET ENVIRONMENT VARIABLE:C812("XCEL_MAIL_RECIPIENT_COUNT"; String:C10($recipients_l))
	
End if 

XCEL_util_EXECUTE("workbook_send_mail_review")