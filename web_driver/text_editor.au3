#include "3rd_party/wd_core.au3"
#include "3rd_party/wd_helper.au3"

Func getTextEditorElement($session)
	return _WD_FindElement($session, $_WD_LOCATOR_ByXPath, "//textarea[@id='textarea']")
EndFunc

Func getSaveButtonElement($session)
	return _WD_FindElement($session, $_WD_LOCATOR_ByXPath, "//img[@id='save']")
EndFunc