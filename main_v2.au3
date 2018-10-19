#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:        Rade & Milos

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <util.au3>
#include <readConfig_v2.au3>

local $scriptFilePath = @ScriptDir&"\config.properties"

local $properties = readConfigFromFile($scriptFilePath)


Global $SEARCH_BOX_X = $properties[0]
Global $SEARCH_BOX_Y = $properties[1]

Global $BROWSER_URL_X = $properties[2]
Global $BROWSER_URL_Y = $properties[3]

Global $SHARE_TEXT_X = $properties[4]
Global $SHARE_TEXT_Y = $properties[5]

Global $SHARE_BUTTON_X = $properties[6]
Global $SHARE_BUTTON_Y = $properties[7]

Global $SHEET1_X = $properties[8]
Global $SHEET1_Y = $properties[9]

Global $SHEET2_X = $properties[10]
Global $SHEET2_Y = $properties[11]

Global $SLEEP_ON_CLOSED_BRACKET = $properties[12]
Global $KEY_ON_CLOSED_BRACKET = $properties[13]

Global $WAIT_FOR_URL_TO_LOAD = $properties[14]

; will dynamically change over time
Global $CURRENT_COPY_TO_ROW = $properties[15]

;########################################
;# TODO
;# Config file [DONE]
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

	showMessage("Done!")
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
	Sleep(500)
EndFunc

Func readFromClipboard()
	$clipboard = ClipGet()
	Send("{ESC}")
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
	second()
	paste()
EndFunc

Func findFirstEmptyCell()
	$current = $CURRENT_COPY_TO_ROW
	While 1
		putInClipboard()
		$value = readFromClipboard()
		If $value = @CRLF Then
			$CURRENT_COPY_TO_ROW = $current
			ExitLoop
		EndIf
		down()
		$current += 1
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

Func second()
	Sleep(800)
EndFunc

Func clearText()
	second()
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
	For $i = 2 To $message[0]-3
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
	Sleep(1000)
	click()
	Sleep(1000)
	space()
EndFunc

Func editSearchBoxFocused()
	click()
	Sleep(1000)
	space()
EndFunc

Func searchCopyToColumn()
	Send("A")
	Send($CURRENT_COPY_TO_ROW)
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
	second()
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