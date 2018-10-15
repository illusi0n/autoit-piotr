#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:        Rade & Milos

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <AutoItConstants.au3>

Local $x
Local $y
Sleep(3000)

leftClick(83,61)

Func leftDoubleClick($x,$y)
MouseClick( $MOUSE_CLICK_LEFT,$x, $y,2)
EndFunc

Func leftClick($x,$y)
MouseClick( $MOUSE_CLICK_LEFT,$x, $y,1)
EndFunc