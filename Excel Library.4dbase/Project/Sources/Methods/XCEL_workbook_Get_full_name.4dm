//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($0)

//equal to short name if not saved to disk yet
SET ENVIRONMENT VARIABLE:C812("XCEL_WORKBOOK_NAME"; $1)

$0:=XCEL_util_EXECUTE("workbook_get_full_name")