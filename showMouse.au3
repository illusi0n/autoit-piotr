#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>

local $x = readConfigFromUI()
showMousePositionNTimes($x)
finishMsg()

Func readConfigFromUI()
	Local $numOfRepeat = InputBox("Configuration", "How many times to repeat operation?", "1", "")
	return $numOfRepeat
EndFunc

Func showMousePositionNTimes($n)
	For $i = 1 To $n
		Sleep(2000)
		showMousePosition()
		Sleep(3000)
	Next
EndFunc
	
Func showMousePosition()
	Local $aPos = MouseGetPos()
	MsgBox($MB_SYSTEMMODAL, "Mouse x, y:", $aPos[0] & ", " & $aPos[1])
EndFunc

Func finishMsg()
	MsgBox($MB_SYSTEMMODAL, "Autoit", "Finished!")
EndFunc