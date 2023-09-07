//%attributes = {"shared":true}
C_TEXT:C284($1)
//  close the workbook without save.
SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)

XCEL_util_EXECUTE("workbook_close")