//%attributes = {"shared":true}
C_REAL:C285($1)
C_REAL:C285($2)
C_REAL:C285($3)
C_REAL:C285($4)

SET ENVIRONMENT VARIABLE:C812("XCEL_LEFT"; String:C10($1))
SET ENVIRONMENT VARIABLE:C812("XCEL_TOP"; String:C10($2))
SET ENVIRONMENT VARIABLE:C812("XCEL_WIDTH"; String:C10($3))
SET ENVIRONMENT VARIABLE:C812("XCEL_HEIGHT"; String:C10($4))

XCEL_util_EXECUTE("application_set_rect")