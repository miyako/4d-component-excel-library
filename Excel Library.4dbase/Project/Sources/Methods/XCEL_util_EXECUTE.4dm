//%attributes = {"invisible":true,"shared":true}
C_TEXT:C284($1)  //name of the script to execute
C_TEXT:C284($0)  //output from the script

_O_PLATFORM PROPERTIES:C365($platform_l)

If ($platform_l=Windows:K25:3)
	
	$0:=VBS_util_EXECUTE($1+".vbs")
	
Else 
	
	$0:=AS_util_EXECUTE($1+".scpt")
	
End if 