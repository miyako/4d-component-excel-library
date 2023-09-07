//%attributes = {"shared":true}
C_TEXT:C284($0)  //name of the new workbook
//by default, "sheet{n}" on Mac, "Book{n}" on PC

//creates a new workbook and returns its name.
//the name can be used to later reference an open workbook.
//on Mac, the application is immediately visible; on PC it is hidden by default.

$0:=XCEL_util_EXECUTE("workbook_add")