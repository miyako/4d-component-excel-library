Function GETENV(variableName)
	
	Set objShell 		= WScript.CreateObject("WScript.Shell")
	Set theVariable 	= objShell.Environment("PROCESS")
	GETENV 			= theVariable(variableName)
	Set objShell 		= Nothing

end Function

Function GETAPP(applicationName)

	On Error Resume Next
	Set GETAPP		 = GetObject(, applicationName)
	If Err.Number <> 0 Then
		Set GETAPP 	= CreateObject(applicationName) 
	End If
	On Error GoTo 0

end Function

Set objExcelApplication	= GETAPP("Excel.Application")
Set theWorkbook			= objExcelApplication.Workbooks(GETENV("XCEL_WORKBOOK_NAME"))
Set theWindow			= theWorkbook.Windows(GETENV("XCEL_WINDOW_CAPTION"))
theDisplayVerticalScrollBarProperty		= GETENV("XCEL_WINDOW_DISPLAY_VERTICAL_SCROLL_BAR")

theWindow.DisplayVerticalScrollBar		= theDisplayVerticalScrollBarProperty
