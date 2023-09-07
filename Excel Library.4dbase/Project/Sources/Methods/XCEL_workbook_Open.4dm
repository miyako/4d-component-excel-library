//%attributes = {"shared":true}
C_TEXT:C284($1)
C_TEXT:C284($0)

SET ENVIRONMENT VARIABLE:C812("XCEL_DOCUMENT_PATH"; $1)

$0:=XCEL_util_EXECUTE("workbook_open")