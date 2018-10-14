#include "3rd_party/wd_core.au3"
#include "3rd_party/wd_helper.au3"
#include "send_keys.au3"
#include "google.au3"
#include "text_editor.au3"

;sendMessageThroughChrome("[Test] Test message", "http://www.mytextarea.com/")
;sendMessageThroughChrome("[Test] Test message", "http://www.google.com/")

Func sendMessageThroughChrome($message, $url)
	;showMessage("[sendMessageThroughChrome] Send message: "&$message)
	;showMessage("[sendMessageThroughChrome] Send to url: "&$url)
	
	Local $capabilities = setupChrome()
	_WD_Startup()

	Local $session = _WD_CreateSession($capabilities)

	openUrl($session, $url)
	Sleep(500)
	$textElement = getTextEditorElement($session)
	click($session, $textElement)
	typeMessage($message)
	sendEscape()
	Sleep(500)
	$saveButton = getSaveButtonElement($session)
	click($session, $saveButton)
	Sleep(500)
	_WD_DeleteSession($session)
	_WD_Shutdown()
EndFunc

$SLEEP_ON_CLOSED_BRACKET = 2000
$KEY_ON_CLOSED_BRACKET = "{ENTER}"

Func typeMessage($stringMessage)
	$message = StringSplit($stringMessage, "")
	For $i = 1 To $message[0]
		sendKey($message[$i])
		If $message[$i] = ']' Then
			Sleep($SLEEP_ON_CLOSED_BRACKET)
			sendKey($KEY_ON_CLOSED_BRACKET)
			Sleep($SLEEP_ON_CLOSED_BRACKET)
		EndIf
	Next
EndFunc

Func openUrl($session, $url)
	_WD_Navigate($session, $url)
EndFunc

Func click($session, $element)
	_WD_ElementAction($session, $element, 'click')
EndFunc

Func setupChrome()
	_WD_Option('Driver', 'chromedriver.exe')
	_WD_Option('Port', 9515)
	_WD_Option('DriverParams', '--log-path="' & @ScriptDir & '\chrome.log"')
	Return '{"capabilities": {"alwaysMatch": {"chromeOptions": {"w3c": true }}}}'
EndFunc
