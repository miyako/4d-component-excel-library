//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($2)
C_BOOLEAN:C305($3)

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_WINDOW_CAPTION"; $2)
SET ENVIRONMENT VARIABLE:C812("XCEL_WINDOW_DISPLAY_HEADINGS"; Choose:C955($3; "True"; "False"))

XCEL_util_EXECUTE("window_set_display_headings")