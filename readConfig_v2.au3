#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <Array.au3>
#include <File.au3>
#include <util.au3>

local $sFilePath = @ScriptDir&"\autoit.config"

local $array = readConfig($sFilePath)
Func readConfig($sFilePath)

;$sReplaceText = "Mani Prakash"
;$sFilePath = "C:\Users\KIRUD01\Desktop\Config.txt"
;$intStartCode = "BuildExe"
$arrRetArray = ""
$s = _FileReadToArray($sFilePath, $arrRetArray);Reading text file and saving it to array $s will show status of reading file..
 For $i = 1 To UBound($arrRetArray)-1
        $line = $arrRetArray[$i];retrieves taskengine text line by line
        If StringInStr($line, "SEAR") Then
                ConsoleWrite ("Starting point "& $line &  @CRLF)
                return StringStripWS(StringSplit($line,":")[2],$STR_STRIPLEADING + $STR_STRIPTRAILING )
        EndIf
        if $i = UBound($arrRetArray)-1 then return "Not Found"
    Next
	return $arrRetArray
EndFunc

_ArrayDisplay($array, "2D display")
showMessage("Done!")
