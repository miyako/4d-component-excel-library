//%attributes = {"invisible":true}
C_TEXT:C284($1; $0)  //script file name; saved as UTF-8

C_LONGINT:C283($platform_l)
_O_PLATFORM PROPERTIES:C365($platform_l)

If ($platform_l=Mac OS:K25:2)
	
	$script_folder_path_t:=Get 4D folder:C485(Current resources folder:K5:16)+"scpt:"
	$script_file_path_t:=$script_folder_path_t+$1
	
	If (Test path name:C476($script_file_path_t)=Is a document:K24:1)
		
		$path_system_t:=Replace string:C233($script_file_path_t; ":"; "/")  //the POSIX separator
		$target_volume_t:=Substring:C12($path_system_t; 1; Position:C15("/"; $path_system_t)-1)
		$system_folder_t:=System folder:C487(System:K41:1)  //take care of the /Volumes/ syntax
		$boot_volume_t:=Substring:C12($system_folder_t; 1; Position:C15(":"; $system_folder_t)-1)
		$script_file_path_t:=Choose:C955($boot_volume_t=$target_volume_t; Substring:C12($path_system_t; Position:C15("/"; $path_system_t)); "/Volumes/"+$path_system_t)
		
		C_BLOB:C604($standard_input_x; $standard_output_x; $standard_error_x)
		LAUNCH EXTERNAL PROCESS:C811("osascript \""+$script_file_path_t+"\""; $standard_input_x; $standard_output_x; $standard_error_x)
		$standard_output_t:=Convert to text:C1012($standard_output_x; "UTF-8")
		
		If (BLOB size:C605($standard_error_x)#0) & (BLOB size:C605($standard_output_x)=0)
			$0:=Convert to text:C1012($standard_error_x; "UTF-8")
		Else 
			$0:=Substring:C12($standard_output_t; 1; Length:C16($standard_output_t)-1)
		End if 
		
	End if 
	
End if 