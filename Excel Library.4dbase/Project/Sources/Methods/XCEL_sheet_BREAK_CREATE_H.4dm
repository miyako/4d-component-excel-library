//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($2)
C_LONGINT:C283($3)

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_SHEET_NAME"; $2)
SET ENVIRONMENT VARIABLE:C812("XCEL_ROW_NUMBER"; String:C10($3))

XCEL_util_EXECUTE("sheet_set_break_horizontal")