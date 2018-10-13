#include "wd_core.au3"
#include "wd_helper.au3"

Local Enum $eFireFox = 0, _
			$eChrome

Local $aDemoSuite[][2] = [["DemoTimeouts", False], ["DemoNavigation", False], ["DemoElements", False], ["DemoScript", False], ["DemoCookies", False], ["DemoAlerts", False],["DemoFrames", False], ["DemoActions", True]]

Local Const $_TestType = $eChrome
Local $sDesiredCapabilities
Local $iIndex
Local $sSession

$_WD_DEBUG = $_WD_DEBUG_Info

Switch $_TestType
	Case $eFireFox
		SetupGecko()

	Case $eChrome
		SetupChrome()

EndSwitch

_WD_Startup()

$sSession = _WD_CreateSession($sDesiredCapabilities)

MyChromeTest()

_WD_DeleteSession($sSession)
_WD_Shutdown()


Func DemoTimeouts()
	_WD_Timeouts($sSession)
	_WD_Timeouts($sSession, '{"pageLoad":2000}')
	_WD_Timeouts($sSession)
EndFunc

Func MyChromeTest()
 
		Local $sElement, $aElements, $sValue

	_WD_Navigate($sSession, "http://google.com")

	
	Sleep(5000)
	$sElement = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//input[@id='lst-ib1']")

	If @error = $_WD_ERROR_NoMatch Then
		$sElement = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//input[@id='lst-ib']")
	EndIf
	
	Sleep(5000)
	_WD_ElementAction($sSession, $sElement, 'value', "testing 123")
	Sleep(5000)
	
	#$sElement = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//input[@name='btnK']")
	#ConsoleWrite('value = ' & $sValue & @CRLF)
	
	$sAction = '{"actions":[{"id":"default mouse","type":"pointer","parameters":{"pointerType":"mouse"},"actions":[{"duration":100,"x":0,"y":0,"type":"pointerMove","origin":{"ELEMENT":"'
	$sAction &= $sElement & '","' & $_WD_ELEMENT_ID & '":"' & $sElement & '"}},{"button":2,"type":"pointerDown"},{"button":2,"type":"pointerUp"}]}]}'

	ConsoleWrite("$sAction = " & $sAction & @CRLF)

	_WD_Action($sSession, "actions", $sAction)
	sleep(5000)
EndFunc

Func DemoElements()
	Local $sElement, $aElements, $sValue

	_WD_Navigate($sSession, "http://google.com")
	$sElement = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//input[@id='lst-ib1']")

	If @error = $_WD_ERROR_NoMatch Then
		$sElement = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//input[@id='lst-ib']")
	EndIf

	$aElements = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//div/input", '', True)

	_ArrayDisplay($aElements)

	_WD_ElementAction($sSession, $sElement, 'value', "testing 123")
	Sleep(500)

	_WD_ElementAction($sSession, $sElement, 'text')
	_WD_ElementAction($sSession, $sElement, 'clear')
	Sleep(500)
	_WD_ElementAction($sSession, $sElement, 'value', "abc xyz")
	Sleep(500)
	_WD_ElementAction($sSession, $sElement, 'text')
	_WD_ElementAction($sSession, $sElement, 'clear')
	Sleep(500)
	_WD_ElementAction($sSession, $sElement, 'value', "fujimo")
	Sleep(500)
	_WD_ElementAction($sSession, $sElement, 'text')
	_WD_ElementAction($sSession, $sElement, 'click')

	_WD_ElementAction($sSession, $sElement, 'Attribute', 'text')

	$sElement = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//input[@id='lst-ib']")
	$sValue = _WD_ElementAction($sSession, $sElement, 'property', 'value')

	ConsoleWrite('value = ' & $sValue & @CRLF)

EndFunc

Func DemoActions()
	Local $sElement, $aElements, $sValue

	_WD_Navigate($sSession, "http://google.com")
	$sElement = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '//input[@id="lst-ib"]')

ConsoleWrite("$sElement = " & $sElement & @CRLF)

	$sAction = '{"actions":[{"id":"default mouse","type":"pointer","parameters":{"pointerType":"mouse"},"actions":[{"duration":100,"x":0,"y":0,"type":"pointerMove","origin":{"ELEMENT":"'
	$sAction &= $sElement & '","' & $_WD_ELEMENT_ID & '":"' & $sElement & '"}},{"button":2,"type":"pointerDown"},{"button":2,"type":"pointerUp"}]}]}'

ConsoleWrite("$sAction = " & $sAction & @CRLF)

	_WD_Action($sSession, "actions", $sAction)
	sleep(5000)
	_WD_Action($sSession, "actions")
	sleep(5000)
EndFunc


Func SetupGecko()
_WD_Option('Driver', 'geckodriver.exe')
_WD_Option('DriverParams', '--log trace')
_WD_Option('Port', 4444)

$sDesiredCapabilities = '{"desiredCapabilities":{"javascriptEnabled":true,"nativeEvents":true,"acceptInsecureCerts":true}}'
EndFunc

Func SetupChrome()
_WD_Option('Driver', 'chromedriver.exe')
_WD_Option('Port', 9515)
_WD_Option('DriverParams', '--log-path="' & @ScriptDir & '\chrome.log"')

$sDesiredCapabilities = '{"capabilities": {"alwaysMatch": {"chromeOptions": {"w3c": true }}}}'
EndFunc
