//%attributes = {"invisible":true}
C_TEXT:C284($1; $0)  //script file name; saved as UTF-16LE
C_BLOB:C604($2)

C_LONGINT:C283($platform_l)
_O_PLATFORM PROPERTIES:C365($platform_l)

If ($platform_l=Windows:K25:3)
	
	$script_folder_path_t:=Get 4D folder:C485(Current resources folder:K5:16)+"vbs\\"
	$script_file_path_t:=$script_folder_path_t+$1
	
	If (Test path name:C476($script_file_path_t)=Is a document:K24:1)
		
		SET ENVIRONMENT VARIABLE:C812("_4D_OPTION_HIDE_CONSOLE"; "true")
		C_BLOB:C604($standard_input_x; $standard_output_x; $standard_error_x)
		
		If (Count parameters:C259=2)
			$standard_input_x:=$2
		End if 
		
		LAUNCH EXTERNAL PROCESS:C811("cscript //Nologo //U \""+$script_file_path_t+"\""; $standard_input_x; $standard_output_x; $standard_error_x)
		$standard_output_t:=Convert to text:C1012($standard_output_x; "UTF-16LE")
		
		If (BLOB size:C605($standard_error_x)#0) & (BLOB size:C605($standard_output_x)=0)
			$0:=Convert to text:C1012($standard_error_x; "UTF-16LE")
		Else 
			$0:=$standard_output_t
		End if 
		
	End if 
	
End if 