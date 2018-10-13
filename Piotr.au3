#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Rade & Milos

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <file.au3>
#include <array.au3>
#include <Excel.au3>
#include <web-driver\wd_core.au3>
#include <web-driver\wd_helper.au3>
#include <web-driver\browserFunctions.au3>

#Local $config = readConfigFromFile("autoit.config")
Local $config = readConfigFromUI()
Local $numOfRepeat = $config[0]
Global $ExcelBook
Local $aResult
$URL = "http://google.com"
$SearchXpath = "//input[@id='lst-ib1']"
executeNTimes($numOfRepeat)
showMessage("Done")

#return array of config property values
#array[0] contains value of how many repeat operations should be
Func readConfigFromUI()
	Local $config[1]
	Local $numOfRepeat = InputBox("Configuration", "How many times to repeat operation?", "1", "")
	$config[0] = $numOfRepeat
	return $config
EndFunc

#return array of config property values
#array[0] contains value of how many repeat operations should be
#file must have the following format
#repeatOperation={value}
Func readConfigFromFile($fileName)
	Dim $allProperties, $i
	$numProperties = _FileReadToArray($fileName, $allProperties)
	Local $config[$allProperties[0]]
	If $numProperties <> 0 Then

		For $i = 1 To $allProperties[0]
			#StringSplit returns array, first element N is size
			#next N elements are array values
			$config[$i-1] = StringSplit($allProperties[$i], "=")[2]
		Next
	EndIf 
	return $config
EndFunc

Func executeNtimes($n)
	For $i = 1 to $n
		showMessage($i)
	Next
EndFunc

Func showArray($array)
	MsgBox(0, "Values", _ArrayToString($array, @TAB))
EndFunc

Func showMessage($message)
	MsgBox(0, "AutoIt", $message)
EndFunc


$FilePath = @ScriptDir &"\excel.xlsx"
$sTextFile = @ScriptDir &"\input.txt"

; running excel
Global $oExcel = _Excel_Open() 
$oExcelBook=_Excel_BookOpen($oExcel, $FilePath)
if @error == 2 Then 
	MsgBox($MB_SYSTEMMODAL, "", "There is no such excel file under the path specified!")
	_Excel_Close($oExcel)
endif

Local $i = 1
$aResult = _Excel_RangeRead($oExcelBook, Default, "F"&$i&":F"&$i, 1)

While Not $aResult = ""
MsgBox($MB_SYSTEMMODAL, "Value", $aResult)
	if $aResult == "NO" Then
		; reading from excel and 
		Local $oInput = _Excel_RangeRead($oExcelBook, Default, "A"&$i&":A"&$i, 1)
		MsgBox($MB_SYSTEMMODAL, "Value", $oInput)
		
		; Splitting string into chars
		$aArray = StringSplit($oInput, "")
		_Excel_Close($oExcel, True, True)
		Sleep(500)
		
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		; Chrome start here, go to Google and click on search bar
		;
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		SetupChrome()
		StartChrome()

		NavigateToURL($URL)

		FindElementByXpath($SearchXpath)
			
		For $elem = 0 to UBound($aArray)-1
			if $aArray[$elem]=="]" Then 
			Sleep(2000)
			SEND("{TAB}")
			Sleep(2000)
			Else 
				SEND($aArray[$elem])
			Endif
			
			Sleep(150)
		Next
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		; Once all text is keyed in, press 'share button'.
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		;8. Go back to excel, copy text from column AF 
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		$oExcel = _Excel_Open()
		$oExcelBook=_Excel_BookOpen($oExcel, $FilePath)
		$oCopy = _Excel_RangeRead($oExcelBook, Default, "AF"&$i&":AF"&$i, 1)
		
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		;9. Change sheet in excel to 2nd sheet
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;		
		
	
	#cs
		$vSheet2 = _ExcelSheetActivate($oExcel, 2)
		$j = 1
		Local $oPaste
		While Not $oPaste ==""
			$oPaste = _Excel_RangeRead($oExcelBook, Default, "A"&$j&":A"&$j, 1)
			if $j==5000 Then 
				showMessage("Quitting after trying to write in cell A"&$j)
				exit 
			Endif
			$j+=1
		WEnd
	#ce
		_ExcelSheetActivate($oExcel, 2)

		$cellNum= giveNumberOfFirstEmptyCell($oExcelBook,"A")
		$FirstEmptyCellInColumn = "A"&$cellNum
		_Excel_RangeWrite ( $oExcelBook, $vSheet2, $oCopy, $vRange = $FirstEmptyCellInColumn , $bValue = True , $bForceFunc = False )
					
	endif
$i += 1
$aResult = _Excel_RangeRead($oExcelBook, Default, "F"&$i&":F"&$i, 1)
WEnd



;If @error Then 
;Exit MsgBox($MB_SYSTEMMODAL, "Error", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
;MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead", "Data successfully read." & @CRLF & "Please click 'OK'")
;endif

Func StartChrome()
	 _WD_Startup()
	$sSession = _WD_CreateSession($sDesiredCapabilities)
EndFunc

Func NavigateToURL($URL)
	_WD_Navigate($sSession, $URL)
EndFunc

Func FindElementByXpath($Xpath)
	$sElement = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, $Xpath)
EndFunc

Func CloseChrome()
	_WD_DeleteSession($sSession)
	_WD_Shutdown()
EndFunc

Func giveValueOfCell($oExcelBook,$Column,$j)
	local $j
	$value = _Excel_RangeRead($oExcelBook, Default, $Column&$j&":"&$Column&$j, 1)			
	Return $value
EndFunc

Func giveNumberOfFirstEmptyCell($oExcelBook,$Column)
local $j
	While Not $oPaste ==""
				$oPaste = _Excel_RangeRead($oExcelBook, Default, $Column&$j&":"&$Column&$j, 1)
				$j+=1
				if $j==5000 Then 
					showMessage("Quitting after trying to write in cell: "&$Column&$j)
					exit 
				Endif
				
	WEnd
Return $j
EndFunc