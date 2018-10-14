#include <file.au3>
#include <array.au3>
#include <util.au3>

#Local $config = readConfigFromFile("autoit.config")
Local $config = readConfigFromUI()
Local $numOfRepeat = $config[0]

executeNTimes($numOfRepeat)
showMessage("Done")

#return array of config property values
#array[0] contains value of how many repeat operations should be
Func readConfigFromUI()
	Local $config[1]
	Local $numOfRepeat = InputBox("Configuration", "How many times to repeat operation?", "1", "")
	$config[0] = $numOfRepeat
	return $config
EndFunc

#return array of config property values
#array[0] contains value of how many repeat operations should be
#file must have the following format
#repeatOperation={value}
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
EndFunc

Func executeNtimes($n)
	For $i = 1 to $n
		showMessage($i)
	Next
EndFunc