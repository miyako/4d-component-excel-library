//%attributes = {"shared":true}
C_TEXT:C284($1)
C_LONGINT:C283($0)

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)

$0:=Num:C11(XCEL_util_EXECUTE("sheet_count"))