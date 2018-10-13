#include "3rd_party/wd_core.au3"
#include "3rd_party/wd_helper.au3"
#include "send_keys.au3"
#include "google.au3"
#######################
#Functions in google.au3
#getSearchButtonElement
#getSearchElement
#######################
#include "text_editor.au3"
#Functions in text_editor.au3
#getTextEditorElement
#getSaveButtonElement
#########################

sendMessageThroughChrome("[Test] Test message", "http://www.mytextarea.com/")
#sendMessageThroughChrome("[Test] Test message", "http://www.google.com/")

Func sendMessageThroughChrome($message, $url)

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

Func typeMessage($stringMessage)
	$message = StringSplit($stringMessage, "")
	For $i = 1 To $message[0]
		sendKey($message[$i])
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
	return '{"capabilities": {"alwaysMatch": {"chromeOptions": {"w3c": true }}}}'
EndFunc
