//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($2)
C_REAL:C285($0)

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_SHEET_NAME"; $2)

$0:=Num:C11(XCEL_util_EXECUTE("setup_get_bottom_margin"))