//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($2)
C_TEXT:C284($3)
C_TEXT:C284($4)

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_SHEET_NAME"; $2)
SET ENVIRONMENT VARIABLE:C812("XCEL_RANGE"; $3)
SET ENVIRONMENT VARIABLE:C812("XCEL_FONT_NAME"; $4)

XCEL_util_EXECUTE("range_set_font_name")