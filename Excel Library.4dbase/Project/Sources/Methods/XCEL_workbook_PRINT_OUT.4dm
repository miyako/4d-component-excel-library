//%attributes = {"shared":true}
C_TEXT:C284($1)
C_LONGINT:C283($2)

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_NUMBER_OF_COPIES"; String:C10($2))

XCEL_util_EXECUTE("workbook_print_out")