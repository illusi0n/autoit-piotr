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

; running excel
Global $oExcel = _Excel_Open() 

$FilePath = @ScriptDir &"\excel.xlsx"
$sTextFile = @ScriptDir &"\input.txt"

$oExcelBook=_Excel_BookOpen($oExcel, $FilePath)
if @error == 2 Then 
	MsgBox($MB_SYSTEMMODAL, "", "There is no such excel file under the path specified!")
	_Excel_Close($oExcel)
endif

Local $i = 1
Local $aResult = _Excel_RangeRead($oExcelBook, Default, "F"&$i&":F"&$i, 1)

While Not $aResult = ""
MsgBox($MB_SYSTEMMODAL, "Value", $aResult)
	if $aResult == "NO" Then
		; reading from excel and 
		Local $oInput = _Excel_RangeRead($oExcelBook, Default, "A"&$i&":A"&$i, 1)
		MsgBox($MB_SYSTEMMODAL, "Value", $oInput)
		
		; Splitting string into chars
		$aArray = StringSplit($oInput, "")
		;_ArrayDisplay($aArray, "", Default, 8)
		_Excel_Close($oExcel, True, True)
		Sleep(500)
		
		; running Notepad and setting it on top
		Run ( "notepad.exe " & $sTextFile, @WindowsDir, @SW_MAXIMIZE )
		
		If WinActivate("[CLASS:Notepad]", "") Then
        MsgBox($MB_SYSTEMMODAL + $MB_ICONWARNING, "Warning", "Window activated" & @CRLF & @CRLF & "May be your system is pretty fast.")
		Else
        ; Notepad will be displayed as MsgBox introduce a delay and allow it.
		MsgBox($MB_SYSTEMMODAL, "", "Window not activated" & @CRLF & @CRLF & "But notepad in background due to MsgBox.", 5)
		EndIf
			
		For $elem = 0 to UBound($aArray)-1
		SEND($elem)
		Sleep(1000)
		Next
		; Close the Notepad window using the classname of Notepad.
		If WinClose("[CLASS:Notepad]", "") Then
			MsgBox($MB_SYSTEMMODAL, "", "Window closed")
		Else
			MsgBox($MB_SYSTEMMODAL + $MB_ICONERROR, "Error", "Window not Found")
		EndIf

	endif
$i += 1
$aResult = _Excel_RangeRead($oExcelBook, Default, "F"&$i&":F"&$i, 1)
WEnd

If @error Then 
Exit MsgBox($MB_SYSTEMMODAL, "Error", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 2", "Data successfully read." & @CRLF & "Please click 'OK' to display the formulas of cells A1:C1 of sheet 2.")
endif

