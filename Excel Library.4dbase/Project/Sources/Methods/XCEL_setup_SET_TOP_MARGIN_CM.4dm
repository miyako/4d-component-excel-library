//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($2)
C_REAL:C285($3)

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_SHEET_NAME"; $2)
SET ENVIRONMENT VARIABLE:C812("XCEL_PAGE_SETUP_TOP_MARGIN"; String:C10($3))

XCEL_util_EXECUTE("setup_set_top_margin_cm")