//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($2)
C_TEXT:C284($3)
C_REAL:C285($4)

SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)
SET ENVIRONMENT VARIABLE:C812("XCEL_SHEET_NAME"; $2)
SET ENVIRONMENT VARIABLE:C812("XCEL_SHAPE_NAME"; $3)
SET ENVIRONMENT VARIABLE:C812("XCEL_SHAPE_ROTATION"; String:C10($4))

XCEL_util_EXECUTE("sheet_shape_set_rotation")