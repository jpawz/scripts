#include <File.au3>
#include <Array.au3>
#include <AutoItConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

Opt("WinTitleMatchMode", -2)

Const $notepadPath = "D:\notepad++\notepad++.exe"
Const $decoderName = "dekoder.exe"
Const $outputDir = "odkodowane"

Prepare()
DecodeFiles()

Func Prepare()
	If not FileExists($notepadPath) Then
		Help()
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

Func Help()
	Local $helpGui = GUICreate("Instrukcja", 300, 300)
	Local $helpEdit = GUICtrlCreateEdit("", 8, 5, 290, 290, BitOR($ES_AUTOVSCROLL, $ES_READONLY, $WS_VSCROLL))
	ConsoleWriteGUI($helpEdit, "Instrukcja:" & @CRLF)
	ConsoleWriteGUI($helpEdit, "1. Notepad++ musi być w katalogu D:\notepad++\" & @CRLF)
	ConsoleWriteGUI($helpEdit, "2. notepad++: Ustawienia -> Ustawienia -> Domyślna ścieżka, zaznaczyć ""Używaj nowego okna..."" "& @CRLF)
	ConsoleWriteGUI($helpEdit, "3. Zaktualizować notepad++ (menu ""?"" -> ""Uaktualij Notepad++"""& @CRLF)
	ConsoleWriteGUI($helpEdit, "4. Zmniejszyć okno Notepad++ (nie powinno być zmaksymalizowane)" & @CRLF)
	ConsoleWriteGUI($helpEdit, "5. Zamknąć Notepad++ i uruchomić odkodowanie jeszcze raz" & @CRLF)
	GUISetState(@SW_SHOW, $helpGui)
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				ExitLoop
		EndSwitch
	Wend
	GUIDelete()
EndFunc

Func ConsoleWriteGUI(Const ByRef $hConsole, Const $sTxt)
    Local Static $sContent = ""
    $sContent &= $sTxt
    GUICtrlSetData($hConsole, $sContent)
EndFunc
