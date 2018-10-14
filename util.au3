#include <array.au3>

Func showMessage($message)
	MsgBox(0, "AutoIt", $message)
EndFunc

Func showArray($array)
	MsgBox(0, "AutoIt", _ArrayToString($array, @TAB))
EndFunc