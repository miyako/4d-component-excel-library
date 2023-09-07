//%attributes = {"invisible":true}
//methods:
//http://msdn.microsoft.com/en-us/library/bb225767(v=office.12).aspx
//properties:
//http://msdn.microsoft.com/en-us/library/bb245609(v=office.12).aspx

XCEL_workbook_close_all
$workbook:=XCEL_workbook_create
XCEL_application_show

//SaveAs
//http://msdn.microsoft.com/en-us/library/bb214129(v=office.12).aspx
$workbook:=XCEL_workbook_save_as_xml($workbook; System folder:C487(Desktop:K41:16)+"sample.xml")
$workbook:=XCEL_workbook_save_as_csv($workbook; System folder:C487(Desktop:K41:16)+"sample.csv")
$workbook:=XCEL_workbook_save_as_sylk($workbook; System folder:C487(Desktop:K41:16)+"sample.slk")
$workbook:=XCEL_workbook_save_as_dif($workbook; System folder:C487(Desktop:K41:16)+"sample.dif")
$workbook:=XCEL_workbook_save_as_xls($workbook; System folder:C487(Desktop:K41:16)+"sample.xls")
$workbook:=XCEL_workbook_save_as_xlsx($workbook; System folder:C487(Desktop:K41:16)+"sample.xlsx")

//note: a newly created workbook must be 'saved as', before it can be 'saved'

//for full list of file formats, see;
//http://msdn.microsoft.com/en-us/library/bb241279(v=office.12).aspx

//if the full name is not a document path, the workbook is probably not saved yet
$full_name_t:=XCEL_workbook_get_full_name($workbook)

//Save
//http://msdn.microsoft.com/en-us/library/bb177993(v=office.12).aspx
XCEL_workbook_save($workbook)

//note: on Mac, the overwrite property is set to True.
//on Windows we suspend alerts to achieve the same effect

//SendForReview
//http://msdn.microsoft.com/en-us/library/bb178022(v=office.12).aspx
XCEL_workbook_send_mail_review($workbook; "miyako@4d-japan.com"; "test")

//SendMail
//http://msdn.microsoft.com/en-us/library/bb178034(v=office.12).aspx
XCEL_workbook_send_mail($workbook; "miyako@4d-japan.com"; "test")

//on Mac, the send mail method opens Mail with the document attatched;
//the Windows equivalent is SendForReview.
//(SendMail onWindows will send immediately)
//on Windows, the workbook must be saved in 'xlShared (2)' mode for it to be attached to am e-mail

//PrintOut
//http://msdn.microsoft.com/en-us/library/bb179158(v=office.12).aspx
XCEL_workbook_print_out($workbook; 1)

//WebPagePreview
//http://msdn.microsoft.com/en-us/library/bb210042(v=office.12).aspx
XCEL_workbook_web_page_preview($workbook)

TRACE:C157

//Close
//http://msdn.microsoft.com/en-us/library/bb179153(v=office.12).aspx
XCEL_workbook_close($workbook)

//note: since we have a 'save' method, 'close' will not save any changes.
//as documented, the auto close macros are not fired in this context.

//test on both Mac and PC