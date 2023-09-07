//%attributes = {"shared":true}
C_BOOLEAN:C305($1)

SET ENVIRONMENT VARIABLE:C812("XCEL_DISPLAY_FULL_SCREEN"; Choose:C955($1; "True"; "False"))

$0:=XCEL_util_EXECUTE("application_set_display_fullscreen")