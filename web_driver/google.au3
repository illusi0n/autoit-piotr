#include "3rd_party/wd_core.au3"
#include "3rd_party/wd_helper.au3"

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