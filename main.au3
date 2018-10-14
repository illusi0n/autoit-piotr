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
#include <web_driver/browser_functions.au3>
#include <util.au3>

#Global variables
$FilePath = @ScriptDir &"\excel.xlsx"
$sTextFile = @ScriptDir &"\input.txt"

$COPY_FROM_COLUMN = "AF"
$COPY_TO_COLUMN = "A"
$STATUS_COLUMN = "AG"
$MESSAGE_COLUMN = "AH"
$URL_COLUMN = "AA"

#Local $config = readConfigFromFile("autoit.config")
Local $config = readConfigFromUI()
Local $numOfRepeat = $config[0]

Local $aResult
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
		executeOnce()
	Next
EndFunc

Func executeOnce()
	; running excel
	$excel = _Excel_Open() 
	$oExcelBook =_Excel_BookOpen($excel, $FilePath)

	if @error == 2 Then 
		MsgBox($MB_SYSTEMMODAL, "", "There is no such excel file under the path specified!")
		_Excel_Close($excel)
	endif

	$i = 1
	$aResult = readCell($oExcelBook, $STATUS_COLUMN&$i, 1)

	While Not $aResult = ""
		MsgBox($MB_SYSTEMMODAL, "Value", $aResult)
			if $aResult == "NO" Then
				; reading from excel and 
				$url = readCell($oExcelBook, $URL_COLUMN&$i, 1)
				$message = readCell($oExcelBook, $MESSAGE_COLUMN&$i, 1)
				MsgBox($MB_SYSTEMMODAL, "Value", $message)
				
				Sleep(500)

				;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
				; Chrome start here, go to $URL (textarea.com)
				; Type text, click save and exit
				;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
				sendMessageThroughChrome($message, $url)

				copyCellValue($oExcelBook, $i)
			endif
		$i += 1
		$aResult = readCell($oExcelBook, $STATUS_COLUMN&$i, 1)
	WEnd
	_Excel_Close($excel, True, True)
EndFunc

;If @error Then 
;Exit MsgBox($MB_SYSTEMMODAL, "Error", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
;MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead", "Data successfully read." & @CRLF & "Please click 'OK'")
;endif

Func getFirstEmptyCell($excelBook)
	$j = 1
	$copyToCell = $COPY_TO_COLUMN&$j
	$cellValue = readCell($excelBook, $COPY_TO_COLUMN&$j, 2)
	While $cellValue <> ""
				$j+=1
				$copyToCell = $COPY_TO_COLUMN&$j
				$cellValue = readCell($excelBook, $copyToCell, 2)
				if $j==5000 Then 
					showMessage("Quitting after trying to write in cell: "&$COPY_TO_COLUMN&$j)
					exit 
				Endif
	WEnd
	Return $copyToCell
EndFunc

Func copyCellValue($excelBook, $copyCellRow)
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	;8. Go back to excel, copy text from column $COPY_FROM_COLUMN
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	$copyFromCell = $COPY_FROM_COLUMN&$copyCellRow
	$copiedValue = readCell($excelBook, $copyFromCell, 1)

	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	;9. Change sheet in excel to 2nd sheet and copy value
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;			
	$copyToCell = getFirstEmptyCell($excelBook)
	
	_Excel_RangeWrite($excelBook, $excelBook.Activesheet, $copiedValue, $copyToCell , True , False )
EndFunc

Func readCell($excelBook, $cell, $sheet)
	$excelBook.Sheets($sheet).Select
	return _Excel_RangeRead($excelBook, Default, $cell&":"&$cell, 1)
EndFunc