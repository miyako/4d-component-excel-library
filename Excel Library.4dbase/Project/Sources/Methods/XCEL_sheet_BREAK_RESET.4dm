//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($2)

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_SHEET_NAME"; $2)

XCEL_util_EXECUTE("sheet_set_break_reset")