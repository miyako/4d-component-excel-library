//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($2)
C_BOOLEAN:C305($0)

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_WINDOW_CAPTION"; $2)

$0:=(XCEL_util_EXECUTE("window_get_display_vertical_scrool_bar")="True")