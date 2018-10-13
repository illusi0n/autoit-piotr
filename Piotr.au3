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

#Local $config = readConfigFromFile("autoit.config")
Local $config = readConfigFromUI()
Local $numOfRepeat = $config[0]

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


Global $oExcel = _Excel_Open() 

$FilePath = @ScriptDir &"\excel.xlsx"

$oExcelBook=_Excel_BookOpen ($oExcel, $FilePath)
if @error == 2 Then 
	MsgBox($MB_SYSTEMMODAL, "", "There is no such excel file under the path specified!")
	_Excel_Close($oExcel)
endif

Local $i = 1
Local $aResult = _Excel_RangeRead($oExcelBook, Default, "F"&$i&":F"&$i, 1)

While Not $aResult = ""
MsgBox($MB_SYSTEMMODAL, "Value", $aResult)
Local $oInput = _Excel_RangeRead($oExcelBook, Default, "A"&$i&":A"&$i, 1)
MsgBox($MB_SYSTEMMODAL, "Value", $oInput)
$i += 1
$aResult = _Excel_RangeRead($oExcelBook, Default, "F"&$i&":F"&$i, 1)
WEnd

If @error Then 
Exit MsgBox($MB_SYSTEMMODAL, "Error", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 2", "Data successfully read." & @CRLF & "Please click 'OK' to display the formulas of cells A1:C1 of sheet 2.")
endif
