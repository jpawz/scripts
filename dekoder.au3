#include <File.au3>
#include <Array.au3>

Opt("WinTitleMatchMode", -2)

Const $notepadPath = "D:\notepad++\notepad++.exe"
Const $decoderName = "dekoder.exe"
Const $outputDir = "odkodowane"

Prepare()
DecodeFiles()

Func Prepare()
	If not FileExists($notepadPath) Then
		MsgBox($MB_OK, "Notepad++", "Nie znaleziono " & $notepadPath)
		Exit
	Endif
	DirCreate($outputDir)
EndFunc

Func DecodeFiles()
	Local $aFileList = _FileListToArray(".", "*", $FLTA_FILES)
	For $i = 1 to UBound($aFileList) - 1
		If StringCompare($aFileList[$i], $decoderName) <> 0 Then
			LaunchN($aFileList[$i])
		Endif
	Next
	MsgBox($MB_OK, "Dekoder", "Zrobione!")
EndFunc


Func LaunchN($aFile)
	Run($notepadPath & " -nosession " & $aFile)
	Local $hWnd = WinWaitActive("[CLASS:Notepad++]")
	ControlSend($hWnd, "Notepad++", "", "!^s")
	Local $sW = WinWaitActive("Zapisywanie jako")
	ControlSend($sW, "[CLASS:#32770]", "[CLASS:Edit; INSTANCE:1]", $outputDir & "\" & $aFile)
	ControlSend($sW, "[CLASS:#32770]", "[CLASS:Edit; INSTANCE:1]", "{ENTER}")
	WinWait($hWnd)
	ControlFocus($hWnd, "Notepad++", "")
	WinClose($hWnd)
EndFunc
