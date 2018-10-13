#include "wd_core.au3"
#include "wd_helper.au3"
#include "send_keys.au3"

Local $sDesiredCapabilities

SetupChrome()

_WD_Startup()

Local $session = _WD_CreateSession($sDesiredCapabilities)
openUrl($session, "http://google.com")
$searchElement = getSearchElement($session)
click($session, $searchElement)
_WD_ElementAction($session, $searchElement, 'value', "testing 123")
sendEscape()
Sleep(500)
$searchButton = getSearchButtonElement($session)
click($session, $searchButton)
Sleep(500)

_WD_DeleteSession($session)
_WD_Shutdown()

Func openUrl($session, $url)
	_WD_Navigate($session, $url)
EndFunc

Func click($session, $element)
	_WD_ElementAction($session, $element, 'click')
EndFunc

Func getSearchElement($session)
	$element = _WD_FindElement($session, $_WD_LOCATOR_ByXPath, "//input[@id='lst-ib1']")

	If @error = $_WD_ERROR_NoMatch Then
		$element = _WD_FindElement($session, $_WD_LOCATOR_ByXPath, "//input[@id='lst-ib']")
	EndIf

	return $element
EndFunc

Func getSearchButtonElement($session)
	return _WD_FindElement($session, $_WD_LOCATOR_ByXPath, "//input[@name='btnK']")
EndFunc

Func SetupChrome()
_WD_Option('Driver', 'chromedriver.exe')
_WD_Option('Port', 9515)
_WD_Option('DriverParams', '--log-path="' & @ScriptDir & '\chrome.log"')

$sDesiredCapabilities = '{"capabilities": {"alwaysMatch": {"chromeOptions": {"w3c": true }}}}'
EndFunc
