//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($2)
C_REAL:C285($3)

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_WINDOW_CAPTION"; $2)
SET ENVIRONMENT VARIABLE:C812("XCEL_WINDOW_TAB_RATIO"; String:C10($3))

XCEL_util_EXECUTE("window_set_tab_ratio")