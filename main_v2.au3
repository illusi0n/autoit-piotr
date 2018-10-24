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
#include <Misc.au3>

local $scriptFilePath = @ScriptDir&"\config2.properties"

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
Global $TIMES_TO_EXEC = $properties[16]
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

Local $numOfExec = $TIMES_TO_EXEC

#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Autoit", 300, 25, 500, 5)
$setMessage = GUICtrlCreateLabel(" Keep pressed ESC to exit!",70,-1)
GUISetState(@SW_SHOW)
WinSetOnTop($Form1, "",1)
#EndRegion ### END Koda GUI section ###

If $numOfExec > 0 Then
	executeNTimes($numOfExec)
Else 
	showMessage("When ready press OK to start!")
	executeScript()
Endif

Func executeNTimes($n)
	showMessage("When ready press X to start! Number of executions: "&$n)
	For $i = 1 To $n
		executeScript()
	Next	
EndFunc
	
Func executeScript()
Sleep(2000)
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

Func closeIfClickedEscape()
    If _IsPressed("1B") Then
			Exit
	EndIf
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
		If $status = "No"&@CRLF or $status = "NO"&@CRLF Then
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
	closeIfClickedEscape()
	copy()
	Sleep(500)
EndFunc

Func readFromClipboard()
	closeIfClickedEscape()
	$clipboard = ClipGet()
	Send("{ESC}")
	return $clipboard
EndFunc

Func processRow($row)
	closeIfClickedEscape()
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
	closeIfClickedEscape()
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
	closeIfClickedEscape()
	ClipPut($value)
	second()
	closeIfClickedEscape()
	paste()
EndFunc

Func findFirstEmptyCell()
	$current = $CURRENT_COPY_TO_ROW
	While 1
		closeIfClickedEscape()
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
	closeIfClickedEscape()
	click()
EndFunc

Func editShareText()
	closeIfClickedEscape()
	click()
EndFunc

Func moveToShareText()
	closeIfClickedEscape()
	moveMouse($SHARE_TEXT_X, $SHARE_TEXT_Y)
EndFunc

Func getCopyCellValue($row)	
	closeIfClickedEscape()
	moveToSearchBox()
	editSearchBox()
	searchCopyCell($row)
	putInClipboard()
	return readFromClipboard()
EndFunc

Func getUrl($row)
	closeIfClickedEscape()
	moveToSearchBox()
	editSearchBoxFocused()
	searchUrlCell($row)
	putInClipboard()
	return readFromClipboard()
EndFunc

Func getMessage($row)
	closeIfClickedEscape()
	moveToSearchBox()
	editSearchBoxFocused()
	searchMessageCell($row)
	putInClipboard()
	return readFromClipboard()
EndFunc

Func searchUrlCell($row)
	closeIfClickedEscape()
	Send("A")
	Send("A")
	closeIfClickedEscape()
	Send($row)
	enter()
	closeIfClickedEscape()
EndFunc

Func searchCopyCell($row)
	closeIfClickedEscape()
	Send("A")
	Send("F")
	closeIfClickedEscape()
	Send($row)
	enter()
	closeIfClickedEscape()
EndFunc

Func searchMessageCell($row)
	closeIfClickedEscape()
	Send("A")
	Send("H")
	closeIfClickedEscape()
	Send($row)
	enter()
	closeIfClickedEscape()
EndFunc

Func second()
	Sleep(800)
	closeIfClickedEscape()
EndFunc

Func clearText()
	closeIfClickedEscape()
	second()
	Send("^a")
	closeIfClickedEscape()
	Send("{BACKSPACE}")
	closeIfClickedEscape()
EndFunc

Func editUrlSearchBox()
	click()
	clearText()
	closeIfClickedEscape()
EndFunc

Func enterUrl($stringUrl)
	closeIfClickedEscape()
	$url = StringSplit($stringUrl, "")
	For $i = 1 To $url[0]
		closeIfClickedEscape()
		Send($url[$i])
	Next
	closeIfClickedEscape()
	enter()
EndFunc

Func waitUrlToLoad()
	Sleep($WAIT_FOR_URL_TO_LOAD)
	closeIfClickedEscape()
EndFunc

Func enterMessage($stringMessage)
	$message = StringSplit($stringMessage, "")
	closeIfClickedEscape()
	; indices start from 2 to lenght-2 to skip " that are
	; carried from excel when copying
	For $i = 2 To $message[0]-3
		closeIfClickedEscape()
		Send($message[$i])
		If $message[$i] = ']' Then
			closeIfClickedEscape()
			Sleep($SLEEP_ON_CLOSED_BRACKET)
			closeIfClickedEscape()
			Send($KEY_ON_CLOSED_BRACKET)
			closeIfClickedEscape()
			Sleep($SLEEP_ON_CLOSED_BRACKET)
			closeIfClickedEscape()
		EndIf
	Next
EndFunc

Func editSearchBox()
	closeIfClickedEscape()
	click()
	Sleep(1000)
	closeIfClickedEscape()
	click()
	Sleep(1000)
	space()
	closeIfClickedEscape()
EndFunc

Func editSearchBoxFocused()
	click()
	Sleep(1000)
	space()
	closeIfClickedEscape()
EndFunc

Func searchCopyToColumn()
	closeIfClickedEscape()
	Send("A")
	Send($CURRENT_COPY_TO_ROW)
	enter()
	closeIfClickedEscape()
EndFunc

Func searchStatusCell($row)
	closeIfClickedEscape()
	Send("A")
	Send("G")
	Send($row)
	enter()
	closeIfClickedEscape()
EndFunc

Func searchStatusColumn()
	closeIfClickedEscape()
	Send("A")
	Send("G")
	Send("1")
	enter()
	closeIfClickedEscape()
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
	closeIfClickedEscape()
	MouseClick($MOUSE_CLICK_LEFT)
	MouseClick($MOUSE_CLICK_LEFT)
EndFunc

Func click()
	closeIfClickedEscape()
	MouseClick($MOUSE_CLICK_LEFT)
EndFunc


Func moveMouse($x, $y)
	closeIfClickedEscape()
	MouseMove($x, $y)
	second()
EndFunc

Func enter()
	closeIfClickedEscape()
	Send("{ENTER}")
EndFunc

Func space()
	closeIfClickedEscape()
	Send("{SPACE}")
EndFunc

Func copy()
	closeIfClickedEscape()
	Send("^c")
EndFunc

Func paste()
	closeIfClickedEscape()
	Send("^v")
EndFunc

Func down()
	closeIfClickedEscape()
	Send("{DOWN}")
EndFunc