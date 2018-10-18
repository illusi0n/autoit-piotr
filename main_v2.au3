#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:        Rade & Milos

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>

;######### MOVE TO CONFIG FILE ###########

Global $SEARCH_BOX_X = 102
Global $SEARCH_BOX_Y = 69

Global $BROWSER_URL_X = 923
Global $BROWSER_URL_Y = 81

Global $SHARE_TEXT_X = 719
Global $SHARE_TEXT_Y = 151

Global $SHARE_BUTTON_X = 719
Global $SHARE_BUTTON_Y = 500

Global $SHEET1_X = 105
Global $SHEET1_Y = 643

Global $SHEET2_X = 156
Global $SHEET2_Y = 651

Global $SLEEP_ON_CLOSED_BRACKET = 2000
Global $KEY_ON_CLOSED_BRACKET = "{ENTER}"

Global $WAIT_FOR_URL_TO_LOAD = 3000

;########################################


;########################################
;# TODO
;# Config file
;# 1) Create a config file with above values empty
;# 2) Config file should be commited as empty file
;# 3) Anyone who uses this could setup these values by own needs
;# 
;# Setup initial values
;# 1) function showMousePositionNTimes can be used for getting position of mouse
;#
;# Issues
;# 1) When message is copied from a cell, it has " at the beginning and in the end.
;#	" should be removed.
;# 2) Cleanup chromedriver and logs, not used anymore [DONE]
;########################################

;showMousePositionNTimes(3)
Sleep(3000)
executeScript()

Func executeScript()
	; 1) move to search box
	moveToSearchBox()
	; 2) Click to Edit Search Box
	editSearchBox()
	; 3) enter 'A', 'G', '1' Enter
	searchStatusColumn()
	; 4) For each cell
	forEachStatusCell()


EndFunc

Func forEachStatusCell()
	local $i = 1
	While 1
		; 5) press CTRL+C
		putInClipboard()
;		6) read from clipboard
		$status = readFromClipboard()
		
		;no more cells			
		If $status = @CRLF Then
			ExitLoop
		EndIf

;		7) check if YES is in clipboard.
		If $status = "NO"&@CRLF Then
			processRow($i)
			moveToCurrentStatusCell($i)
		EndIf
		down()
		$i += 1
	WEnd
EndFunc

Func moveToCurrentStatusCell($row)
	moveToSearchBox()
	editSearchBoxFocused()
	searchStatusCell($row)
EndFunc

Func putInClipboard()
	copy()
	oneSecond()
EndFunc

Func readFromClipboard()
	$clipboard = ClipGet()
	oneSecond()
	Send("{ESC}")
	oneSecond()
	return $clipboard
EndFunc

Func processRow($row)
	; 8) get url from AA
	$url = getUrl($row)
	; 9) get message from AH
	$message = getMessage($row)
	
	; 10) move to web url
	moveToWebUrl()
	; 11) edit url
	editUrlSearchBox()
	; 12) search url
	enterUrl($url)
	; 13) load url
	waitUrlToLoad()
	; 14) move to share (demo: textarea) text
	moveToShareText()
	; 15) edit share (demo: textarea
	editShareText()
	; 16) type message
	enterMessage($message)
	; 17) clear message (due to saving previous messages, clear it) #ONLY_FOR_DEMO
	clearText()
	; 18) move to share (demo: save) button
	moveToShareButton()
	; 19) click on save button
	clickOnShareButton()
	; done with browser go to excel to copy cell
	copyCell($row)
EndFunc

Func copyCell($row)
	; 20) copy column cell
	$copiedValue = getCopyCellValue($row)
	; 21) go to sheet2
	moveToSheet2()
	click()
	; 22) find the copy to column
	moveToSearchBox()
	editSearchBoxFocused()
	searchCopyToColumn()
	; 23) move to first empty cell in column
	findFirstEmptyCell()
	; 24) paste copied value to empty cell
	pasteCopiedValue($copiedValue)
	; 25) return to sheet1
	moveToSheet1()
	click()
EndFunc

Func pasteCopiedValue($value)
	ClipPut($value)
	oneSecond()
	paste()
EndFunc

Func findFirstEmptyCell()
	While 1
		putInClipboard()
		$value = readFromClipboard()
		If $value = @CRLF Then
			ExitLoop
		EndIf
		down()
	WEnd
EndFunc

Func clickOnShareButton()
	click()
EndFunc

Func editShareText()
	click()
EndFunc

Func moveToShareText()
	moveMouse($SHARE_TEXT_X, $SHARE_TEXT_Y)
EndFunc

Func getCopyCellValue($row)
	moveToSearchBox()
	editSearchBox()
	searchCopyCell($row)
	putInClipboard()
	return readFromClipboard()
EndFunc

Func getUrl($row)
	moveToSearchBox()
	editSearchBoxFocused()
	searchUrlCell($row)
	putInClipboard()
	return readFromClipboard()
EndFunc

Func getMessage($row)
	moveToSearchBox()
	editSearchBoxFocused()
	searchMessageCell($row)
	putInClipboard()
	return readFromClipboard()
EndFunc

Func searchUrlCell($row)
	Send("A")
	Send("A")
	Send($row)
	enter()
EndFunc

Func searchCopyCell($row)
	Send("A")
	Send("F")
	Send($row)
	enter()
EndFunc

Func searchMessageCell($row)
	Send("A")
	Send("H")
	Send($row)
	enter()
EndFunc

Func oneSecond()
	Sleep(1000)
EndFunc

Func clearText()
	oneSecond()
	Send("^a")
	Send("{BACKSPACE}")
EndFunc

Func editUrlSearchBox()
	click()
	clearText()
EndFunc

Func enterUrl($stringUrl)
	$url = StringSplit($stringUrl, "")
	For $i = 1 To $url[0]
		Send($url[$i])
	Next
	enter()
EndFunc

Func waitUrlToLoad()
	Sleep($WAIT_FOR_URL_TO_LOAD)
EndFunc

Func enterMessage($stringMessage)
	$message = StringSplit($stringMessage, "")
	
	; indices start from 2 to lenght-2 to skip " that are
	; carried from excel when copying
	For $i = 2 To $message[0]-2
		Send($message[$i])
		If $message[$i] = ']' Then
			Sleep($SLEEP_ON_CLOSED_BRACKET)
			Send($KEY_ON_CLOSED_BRACKET)
			Sleep($SLEEP_ON_CLOSED_BRACKET)
		EndIf
	Next
EndFunc

Func editSearchBox()
	click()
	oneSecond()
	click()
	oneSecond()
	space()
EndFunc

Func editSearchBoxFocused()
	click()
	oneSecond()
	space()
EndFunc

Func searchCopyToColumn()
	Send("A")
	Send("1")
	enter()
EndFunc

Func searchStatusCell($row)
	Send("A")
	Send("G")
	Send($row)
	enter()
EndFunc

Func searchStatusColumn()
	Send("A")
	Send("G")
	Send("1")
	enter()
EndFunc

Func moveToSheet1()
	moveMouse($SHEET1_X , $SHEET1_Y)
EndFunc

Func moveToSheet2()
	moveMouse($SHEET2_X , $SHEET2_Y)
EndFunc

Func moveToShareButton()
	moveMouse($SHARE_BUTTON_X, $SHARE_BUTTON_Y)
EndFunc

Func moveToWebUrl()
	moveMouse($BROWSER_URL_X, $BROWSER_URL_Y)
EndFunc

Func moveToSearchBox()
	moveMouse($SEARCH_BOX_X, $SEARCH_BOX_Y)
EndFunc

Func doubleClick()
	MouseClick($MOUSE_CLICK_LEFT)
	MouseClick($MOUSE_CLICK_LEFT)
EndFunc

Func click()
	MouseClick($MOUSE_CLICK_LEFT)
EndFunc


Func moveMouse($x, $y)
	MouseMove($x, $y)
	Sleep(5000)
EndFunc

Func enter()
	Send("{ENTER}")
EndFunc

Func space()
	Send("{SPACE}")
EndFunc

Func copy()
	Send("^c")
EndFunc

Func paste()
	Send("^v")
EndFunc

Func down()
	Send("{DOWN}")
EndFunc

Func showMousePositionNTimes($n)
	For $i = 1 To $n
		showMousePosition()
		Sleep(5000)
	Next
EndFunc
	
Func showMousePosition()
	Local $aPos = MouseGetPos()
	MsgBox($MB_SYSTEMMODAL, "Mouse x, y:", $aPos[0] & ", " & $aPos[1])
EndFunc

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
EndFuncs