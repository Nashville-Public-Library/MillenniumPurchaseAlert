; MILLENNIUM_PURCHASE_ALERT.au3
; JAMES STAUB, NASHVILLE PUBLIC LIBRARY
; 20130606

; Using Millennium and Excel, collect and calculate holds, current holdings, 
; and current orders 'cause Millennium is too stupid to do this itself.
; The final result: a spreadsheet that acquisitions staff can use to drive
; additional copy purchases.

; Caveat : 
; I believe the real goal is to reduce the wait time of patrons; holds to items ratio
; is only an approximation to begin with

; FYI this report is practically real-time. The Lists and High-Demand Holds
; report it is based on appear to be real time and do not appear to be 
; overnight cron jobs

; TO DO
; DETERMINE whether holds ratio OR TIME TO HOLDSHELF is the important variable...
; MAYBE allow configurable Boolean search strategies, e.g., incorporate something like http://research.ahml.info/oml/au3LoadSearchStrategyFile.txt - explained at http://research.ahml.info/oml/AutoIt.html (Find "LoadSearchStrategyFile")
; SORT AND FILTER according to Julie, Laurie, and Phyllis' needs
; MAYBE eliminate the Excel AutoIT UDF stuff in favor of COM

; CONFIGURATION
AutoItSetOption("MouseCoordMode",0)
AutoItSetOption("SendCapslockMode", 0)
AutoItSetOption("SendKeyDelay", 100)
Send("{CAPSLOCK off}")

; VARIABLES
Local $sMillenniumUsername = ""
Local $sMillenniumUsernamePassword = ""
Local $sMillenniumInitials = ""
Local $sMillenniumInitialsPassword = ""
Local $sName = "JAMES-PURCHASEALERT" ; NB: no spaces!
Local $sPath = "C:\Users\jstaub\Desktop\PURCHASE_ALERT\" ; Where to save files, must contain final backslash
Local $sPathSharePoint = "http://intranet/jstaub/Shared Documents/" ; Where to download and save final spreadsheet, must contain final frontslash
Local $sHighDemandHoldsFilename = $sName & "-HIGH-DEMAND_HOLDS" ; NB: no spaces!
Local $nReviewFileBibliographic = 48
Local $sReviewFileNameBibliographic = $sName & "-BIB" ; NB: no spaces!
Local $sReviewFileQueryBibliographic = "{TAB}b{TAB 4}b{TAB}8{TAB}e{TAB}!s"
Local $sReviewFileExportFieldsBibliographic = "b{TAB}81{TAB}!ab{TAB}30{TAB}!ab{TAB}c{TAB}!ab{TAB}a{TAB}!ab{TAB}t{TAB}!ab{TAB}i{TAB}!ab{TAB}{!}{TAB}260|c!o{TAB}!ab{TAB}{!}{TAB}264|c!o{TAB}"
Local $nReviewFileItem = 64
Local $sReviewFileNameItem = $sName & "-ITEM" ; NB: no spaces!
Local $sReviewFileQueryItem = "{TAB}i{TAB}rev{TAB}" & $sReviewFileNameBibliographic & "{TAB}i{TAB}88{TAB}r{TAB}[-{!}cdft]{TAB}!s"
Local $sReviewFileExportFieldsItem = "b{TAB}81{TAB}!ai{TAB}81{TAB}!ai{TAB}83{TAB}!ai{TAB}8{TAB}"
Local $nReviewFileOrder = 64
Local $sReviewFileNameOrder = $sName & "-ORDER" ; NB: no spaces!
Local $sReviewFileQueryOrderCDATEDifference = -7
Local $sReviewFileQueryOrderCDATEStart = _DateAdd('d', $sReviewFileQueryOrderCDATEDifference, _NowCalcDate())
Local $sReviewFileQueryOrder = "{TAB}o{TAB}rev{TAB}" & $sReviewFileNameBibliographic & "{TAB}o{TAB}20{TAB}R{TAB}[1aoq]{TAB}!ao{TAB}03{TAB}={TAB 2}!a+{TAB}o{TAB}o{TAB}03{TAB}W{TAB}" & StringMid($sReviewFileQueryOrderCDATEStart,6,2) & StringMid($sReviewFileQueryOrderCDATEStart,9,2) & StringMid($sReviewFileQueryOrderCDATEStart,3,2) & "{TAB}t^{END}+{UP}!g!s"
Local $sReviewFileExportFieldsOrder = "b{TAB}81{TAB}!ao{TAB}81{TAB}!ao{TAB}05{TAB}!ao{TAB}03{TAB}"

; INCLUDES
#include <Array.au3>
#include <Date.au3>
#include <Excel.au3>

; FUNCTIONS
Func ClickAt( $x, $y )
	MouseClick( "left", $x, $y, 1, 0 )
	Sleep(400)  ;add pause after clicking mouse
EndFunc

Func ExcelImportData($ImportFilename)
   Send("!aft") ; Select Data > Import data From Text
   WinWaitActive(" Import Text File")
   Send("!n") ; Select File name
   ClipPut($sPath & $ImportFilename & ".txt")
   Send("^v")
   Send("!o") ; Select Open
   WinWaitActive("Text Import Wizard - Step 1 of 3")
   Send("!d{SPACE}") ; Select Delimited
   Send("!r1") ; Select Start import at row 1 ; High-Demand Holds maybe start at 3...
   Send("!o{SPACE}6500{DOWN}{ENTER}") ; Select File origin: UTF-8
   Send("!n") ; Select Next
   WinWaitActive("Text Import Wizard - Step 2 of 3")
   Send("!t{+}!m{-}!c{-}!s{-}") ; set Delimiter to Tab (and only Tab) ; Send("!o{-}") ; wants to place a Hyphen in the Other Character textbox, not James' intention!
   Send("!q{HOME}{ENTER}") ; set Text Qualifier to [double quotation mark]. Yikes! How do I send a double quotation mark in AutoIT? ; FYI NOT REALLY APPROPRIATE FOR THE NO-QUOTE CREATE LIST EXPORTS...
   Send("!f") ; Select Finish
   WinWaitActive("Import Data")
   Send("{ENTER}")
   WinWaitActive("Microsoft Excel - ")   
EndFunc

Func ExcelWriteNewColumn($sColumnLabel,$sColumnLetter,$sColumnFormula)
   Local $nRows = $oExcel.ActiveSheet.UsedRange.Rows.Count ; count the ROWS in the Active Sheet for use in later AutoFill commands
   _ExcelWriteCell($oExcel,$sColumnLabel,$sColumnLetter & "1")
   _ExcelWriteFormula($oExcel,$sColumnFormula,$sColumnLetter & "2")
   $oExcel.ActiveSheet.Range($sColumnLetter & "2:" & $sColumnLetter & "2").AutoFill($oExcel.ActiveSheet.Range($sColumnLetter & "2:" & $sColumnLetter & $nRows))
   $oExcel.ActiveSheet.Range($sColumnLetter & "2:" & $sColumnLetter & $nRows).Select
   Send("^c") ; Copy formula values [JAMES should look at VBA methods!]
   Sleep(500)
   Send("!h")
   Sleep(500)
   Send("v")
   Sleep(500)
   Send("v") ; paste as values in place
   Sleep(500)
EndFunc

Func GetHighlightedData()  ; [Harvey Hahn 13 May 2005]
	Dim $HighlightedData
	ClipPut("")  ; clear the clipboard
	Send("^c")  ; copy highlighted data to the clipboard
	Sleep(200)
	Send("^c")  ; duplicate command needed for reliability (don't know why)
	Sleep(200)
	$HighlightedData = ClipGet()  ; put the text on the clipboard into a variable
	Return $HighlightedData
EndFunc
 
Func MillenniumCreateListSearchAndExport($nFile, $sName, $sQuery, $sExport)
   WinActivate("Millennium Administration")
   WinWaitActive("Millennium Administration")
   Send("!g") ; Select Go menu
   Send("l") ; Select Go > Create Lists
   ClickAt(144,238) ; Click on Review File number 1
   Send("^{HOME}") ; Make sure Review File number 1 is selected (in case there's some scrollin' goin' on)
   Send("{DOWN " & $nFile - 1 & "}") ; SELECT THE REVIEW FILE
   ; JAMES SHOULD VERIFY THAT THE SELECTED LIST IS THE DESIRED LIST
   Send("!te") ; Tools > Empty
   Sleep(500)
   If WinExists("Warning") Then 
	  Send("!y") ; Yes Are you sure you want to empty [FILENAME]
   EndIf
   Send("!s") ; Select Search Records
   Sleep(500)
   If WinExists("Warning") Then 
	  Send("!y") ; Yes Do you want to remove the file being copied from
   EndIf
   WinWaitActive("Boolean Search")
   WinActivate("Boolean Search")
   SendLong($sName)
   Send($sQuery)
   WinWaitActive("Search")
   Send("!y") ; Select Yes
   WinActivate("Millennium Administration")
   WinWaitActive("Millennium Administration")
   ClickAt(144,238) ; Click on Review File number 1
   Send("^{HOME}") ; Make sure Review File number 1 is selected (in case there's some scrollin' goin' on)
   Send("{DOWN " & $nFile - 1 & "}") ; SELECT THE REVIEW FILE
   Local $sReviewFileComplete
   Do
	  $sReviewFileComplete = GetHighlightedData()
	  Sleep(2000)
   Until StringInStr($sReviewFileComplete, "complete")
   Send("!x") ; Select Export Records
   WinWaitActive($sName)
   WinActivate($sName)
   Send($sExport)
   Send("!e{SPACE}9!o{TAB}{SPACE}!n!o{TAB}{SPACE}!a{TAB 2}|!o!f") ; Review file export delimiters
   SendLong($sPath & $sName & ".txt") ; Save file name
   Send("!o")
   WinWaitActive("Millennium Administration")
   WinActivate("Millennium Administration")
EndFunc 

Func SendLong($sSendLong)
   ClipPut($sSendLong)
   Send("^v")
EndFunc

; SCRIPT

; MILLENNIUM OPEN
Run("C:\Millennium\iiirunner.exe")
; USER : ENTER USERNAME AND PASSWORD
WinWaitActive("Enter Login and Password")
Send($sMillenniumUsername & "{TAB}" & $sMillenniumUsernamePassword & "{TAB}{ENTER}")
; USER : ENTER INITIALS AND PASSWORD
WinWaitActive("Enter Initials and Password")
Send($sMillenniumInitials & "{TAB}" & $sMillenniumInitialsPassword & "{TAB}{ENTER}")

; CREATE LISTS FOR BIBS, ITEMS, ORDERS, AND EXPORT DATA
MillenniumCreateListSearchAndExport($nReviewFileBibliographic,$sReviewFileNameBibliographic,$sReviewFileQueryBibliographic,$sReviewFileExportFieldsBibliographic)
MillenniumCreateListSearchAndExport($nReviewFileItem,$sReviewFileNameItem,$sReviewFileQueryItem,$sReviewFileExportFieldsItem)
MillenniumCreateListSearchAndExport($nReviewFileOrder,$sReviewFileNameOrder,$sReviewFileQueryOrder,$sReviewFileExportFieldsOrder)

; RUN AND EXPORT HIGH-DEMAND HOLDS REPORT
WinWaitActive("Millennium Administration")
WinActivate("Millennium Administration")
Send("!gc") ; Select Go > Millennium Control Bar
WinWaitActive("Millennium Control Bar")
; MILLENNIUM : OPEN CIRCULATION
WinActivate("Millennium Control Bar")
ClickAt(21,76) ; Click on the Circulation mode icon
WinWaitActive("Enter Initials and Password")
Send($sMillenniumInitials & "{TAB}" & $sMillenniumInitialsPassword & "!o")
WinWaitActive("Due Slip Printing")
Send("!n") ; Select No [due slip printing] button
WinWaitActive("Millennium Circulation")
WinActivate("Millennium Circulation")
Send("!g{DOWN 4}{ENTER}") ; Select Go > High-Demand Holds
Send("!c") ; Select Create Report
Do
   Sleep(1000)
Until WinActive("Millennium Circulation")
Send("!te") ; Select Tools > Export
WinWaitActive("Export Table")
Send("!t") ; Select Sae as tabbed text
WinWaitActive("Enter or select the export output file")
Send("!n") ; Focus on File Name
SendLong($sPath & $sHighDemandHoldsFileName & ".txt")
Send("{ENTER}")

; MILLENNIUM : CLOSE MILLENNIUM CIRCULATION
WinActivate("Millennium Circulation")
WinWaitActive("Millennium Circulation")
WinClose("Millennium Circulation")
WinWaitActive("Question")
Send("!y")
Sleep(1000)
; MILLENNIUM : CLOSE MILLENNIUM CONTROL BAR
WinActivate("Millennium Control Bar")
WinWaitActive("Millennium Control Bar")
WinClose("Millennium Control Bar")
WinWaitActive("Question")
Send("!y")
Sleep(1000)
; MILLENNIUM : CLOSE MILLENNIUM ADMINISTRATION
WinActivate("Millennium Administration")
WinWaitActive("Millennium Administration")
WinClose("Millennium Administration")
WinWaitActive("Question")
Send("!y")
Sleep(1000)

; EXCEL
; Scripting in AutoIT rather than VBA for the simple hell of it.

Local $oExcel = ObjCreate("Excel.Application") ; create excel object
$oExcel.Visible = True
$oExcel.Workbooks.Open($sPathSharePoint & $sName & ".xlsx")

Sleep(1000)
WinSetState("Microsoft Excel - " & $sName,"",@SW_RESTORE)
Sleep(1000)

; PLACE NEW INFORMATION IN ACTION WORKSHEET
_ExcelSheetDelete($oExcel, "ACTION")
_ExcelSheetAddNew($oExcel, "ACTION")
_ExcelSheetActivate($oExcel, "HDH")
$oExcel.ActiveSheet.ShowAllData
$oExcel.ActiveSheet.Range("A:O").AutoFilter(15,"<>")
$oExcel.ActiveSheet.Range("A1").CurrentRegion.Copy
_ExcelSheetActivate($oExcel, "ACTION")
;$oExcel.ActiveSheet.Range("A1").PasteSpecial(xlPasteValues) ; can't figure out the syntax for paste values right now...
Send("^{HOME}")
Send("!h")
Sleep(500)
Send("v")
Sleep(500)
Send("v") ; paste as values in place
Sleep(500)
_ExcelColumnDelete($oExcel, 2, 13) ; Delete Column B - N

_ExcelSheetDelete($oExcel, "ORDER")
_ExcelSheetDelete($oExcel, "ITEM")
_ExcelSheetDelete($oExcel, "BIB")
_ExcelSheetDelete($oExcel, "HDH")

_ExcelSheetAddNew($oExcel, "ORDER")
_ExcelSheetAddNew($oExcel, "ITEM")
_ExcelSheetAddNew($oExcel, "BIB")
_ExcelSheetAddNew($oExcel, "HDH")
; _ExcelSheetDelete($oExcel, "PLACEHOLDER")

_ExcelSheetActivate($oExcel, "BIB")
ExcelImportData($sReviewFileNameBibliographic)

_ExcelSheetActivate($oExcel, "ITEM")
ExcelImportData($sReviewFileNameItem)
; write HOLDSCOUNT(ITEM) column
ExcelWriteNewColumn("HOLDSCOUNT(ITEM)","E",'=(LEN(D2)-LEN(SUBSTITUTE(D2,"P#=","")))/3')
; SORT by CREATED(ITEM)
Send("^{HOME}")
$oExcel.ActiveSheet.Range("A2:E" & $oExcel.ActiveSheet.UsedRange.Rows.Count).Sort($oExcel.ActiveSheet.Range("A2"),1,$oExcel.ActiveSheet.Range("C2"),Default,2)

_ExcelSheetActivate($oExcel, "ORDER")
ExcelImportData($sReviewFileNameOrder)
; write NEWER column [i.e., should we count the ORDER COPIES, or are there items created more recently than the ORDER CDATE?]
ExcelWriteNewColumn("NEWER","E",'=IF(D2>IFERROR(VLOOKUP(A2,ITEM!A:C,3,FALSE),0),"ORDER","ITEM")')
   
_ExcelSheetActivate($oExcel, "HDH")
ExcelImportData($sHighDemandHoldsFilename)

; set column width
$oExcel.Worksheets("HDH").Columns("A:F").ColumnWidth = 10

_ExcelRowDelete($oExcel, 1, 2) ; Delete Rows 1 and 2 of the High-Demand Holds report
_ExcelColumnDelete($oExcel, 1, 1) ; Delete Column 1 - # - of the High-Demand Holds report
_ExcelColumnDelete($oExcel, 6, 1) ; Delete Column 6 - System Items - of the High-Demand Holds report
_ExcelColumnInsert($oExcel, 5, 4) ; Insert 4 Column at column E

WinActivate("Microsoft Excel - " & $sName)
WinWaitActive("Microsoft Excel - " & $sName)

; Embolden column headers
$oExcel.ActiveSheet.Rows("1:1").Select
Send("^b")
; Freeze top row
$oExcel.ActiveSheet.Rows("2:2").Select
$oExcel.ActiveWindow.FreezePanes = True
Send("^{HOME}")

; write CALL NUMBER column
ExcelWriteNewColumn("CALL NUMBER","E",'=IFERROR(VLOOKUP(A2,BIB!A:C,3,FALSE),"")')
; write JUV? column
ExcelWriteNewColumn("AUDIENCE","F",'=IF(OR(LEFT(E2,1)="j",LEFT(E2,1)="e"),"JUV",IF(LEFT(E2,2)="ya","YA","ADULT"))')
; write PUB DATE column using 260|c or 264|c
ExcelWriteNewColumn("PUB DATE","G",'=IFERROR(VLOOKUP(A2,BIB!A:H,7,FALSE),IFERROR(VLOOKUP(A2,BIB!A:H,8,FALSE),""))')
; write ISN column
ExcelWriteNewColumn("ISN","H",'=IFERROR(TEXT(VLOOKUP(A2,BIB!A:F,6,FALSE),"#"),"")')
; write BIB HOLDS column [Really - BIB holds + item records with 2 or more holds that should be re-assigned]
ExcelWriteNewColumn("BIB HOLDS","J",'=I2-COUNTIFS(ITEM!A:A,A2,ITEM!C:C,">0")')
; Delete System Holds column
_ExcelColumnDelete($oExcel, 9, 1) ; Delete Column 1 - # - of the High-Demand Holds report
; write VALID ITEMS column
ExcelWriteNewColumn("VALID ITEMS","J",'=COUNTIF(ITEM!A:A,A2)')
; write ON ORDER COPIES column
ExcelWriteNewColumn("ON ORDER COPIES","K",'=SUMIFS(ORDER!C:C,ORDER!A:A,HDH!A2,ORDER!E:E,"ORDER")')
; write RATIO column - JAMES MAYBE CHANGE TO INCLUDE "GONE" STUFF AS A NEGATIVE RATIO???
ExcelWriteNewColumn("RATIO","L",'=IFERROR(I2/(J2+K2),-I2)')
; write ORDER DECISION column - JAMES MAYBE CHANGE TO INCLUDE "GONE" STUFF NEGATIVE RATIO IN A BETTER WAY?
ExcelWriteNewColumn("ORDER DECISION","M",'=IF(E2="","0 BIB HOLDS",IF(L2<0,"ORDER",IF(L2>(IF(D2="DVD",8,5)),"ORDER","DO NOT ORDER")))')
; write ORDER COPIES column
ExcelWriteNewColumn("ORDER COPIES","N",'=IF(L2<0,ROUNDUP(I2/(IF(D2="DVD",8,5)),0),IF(L2>(IF(D2="DVD",8,5)),ROUNDUP((I2/(IF(D2="DVD",8,5)))-(J2+K2),0),""))')
; write ACTION column
ExcelWriteNewColumn("ACTION","O",'=IFERROR(VLOOKUP(A2,ACTION!A:C,2,FALSE),"")')

; SORT by CALL Number
Send("^{HOME}")
$oExcel.ActiveSheet.Range("A2:O" & $oExcel.ActiveSheet.UsedRange.Rows.Count).Sort($oExcel.ActiveSheet.Range("E2"),1)
; AutoFilter
$oExcel.ActiveSheet.Range("A:O").AutoFilter(13,"ORDER")
$oExcel.ActiveSheet.Range("A:O").AutoFilter(12,">0")
$oExcel.ActiveSheet.Range("A:O").AutoFilter(15,"")

; Save Excel Spreadsheet
$oExcel.Workbooks($sName & ".xlsx").Save
; Save Excel Spreadsheet to SharePoint
$oExcel.Workbooks($sName & ".xlsx").SaveAs($sPathSharePoint & $sName)
; Close Excel Spreadsheet
$oExcel.Workbooks($sName & ".xlsx").Close
; Close Excel window
WinClose("Microsoft Excel")

Exit

#cs ----------------------------------------------------------------------------
Bibliographic search
+++++
Nashville
JAMES-PURCHASEALERT-BIB
b

1			b	8	e			

BIBLIOGRAPHIC  HOLD  exist      

+++++
Multnomah County

1			b	8	e			
2	A		b	!599	A	out of print		
3	A		b	31	!=			
4	A		b	81	!=	1332990		

BIBLIOGRAPHIC  HOLD  exist      AND BIBLIOGRAPHIC  MARC Tag 599  All Fields don't have  "out of print"    AND BIBLIOGRAPHIC  BCODE3  not equal to  ""    AND BIBLIOGRAPHIC  RECORD #  not equal to  "1332990"

#ce ----------------------------------------------------------------------------

#cs ----------------------------------------------------------------------------
Item search based on Review File created by Bibliographic search
+++++
Nashville
[perhaps expand to include ASAA Rules for Requesting, e.g., to exclude Nashville Room stuff... but wait, those should have Library Use Only statuses...]

1			i	88	=	-		
2	O		i	88	=	!		
3	O		i	88	=	c		
4	O		i	88	=	d		
5	O		i	88	=	f		
6	O		i	88	=	t		

[easier way to write this is:]
1			i	88	R	[-!cdft]		

+++++
Multnomah County

1			i	61	!=	135		
2	A		i	61	N	190	192	
3	A	(	i	88	=	t		
4	O		i	88	=	x		
5	O		i	88	=	!		
6	O		i	88	=	p		
7	O		i	88	=	-		)
8	A		i	94	=	0		
9	A		i	97	!=	f		

ITEM  I TYPE  not equal to  "135"    AND ITEM  I TYPE  not within  "190"and "192"    AND (ITEM  STATUS  equal to  "t"    OR ITEM  STATUS  equal to  "x"    OR ITEM  STATUS  equal to  "!"    OR ITEM  STATUS  equal to  "p"    OR ITEM  STATUS  equal to  "-")    AND ITEM  DON'T USE  equal to  "0"    AND ITEM  IMESSAGE  not equal to  "f"

#ce ----------------------------------------------------------------------------

#cs ----------------------------------------------------------------------------
Order search based on Review File created by Bibliographic search
+++++
Nashville
NPL does not have many - if any - partially paid orders. Exclude from search.
NPL has maybe 50-ish order records with STATUS = ON ORDER and an RDATE value more than 1 day in the past. JAMES should have Melissa Meyers look at these.
NPL will use ORDER CAT DATE (03) to determine whether order record COPIES should be included. If ORDER CAT DATE does not exist, then chances are ITEM records have not been created yet. If ORDER CAT DATE does exist, chances are high that item records have been created - or will be in the very near future!

1			o	20	R	[1aoq]		
2	A	(	o	03	=	  -  -  	  -  -  	
3	O		o	03	W	[today - 7]	[today]	) ; MAYBE add this last tweak...

1			o	20	R	[1aoq]		
2	A	(	o	03	=	  -  -  	  -  -  	
3	O		o	03	W	05-16-2013	05-23-2013	)

+++++
Multnomah County

1		(	o	20	=	o		
2	O		o	20	=	q		)
3	A		o	12	!=	pald		
4	A		o	17	=	  -  -  	  -  -  	

(ORDER  STATUS  equal to  "o"    OR ORDER  STATUS  equal to  "q")    AND ORDER  FUND  not equal to  "pald"    AND ORDER  RDATE  equal to  "  -  -  "

#ce ----------------------------------------------------------------------------
