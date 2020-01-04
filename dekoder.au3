; TODO: 
; * progress bar https://www.autoitscript.com/wiki/Progress_Bar_Sample
; * save instead of save as (faster) - overwrite existing files

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
	Local $aFileList = _FileListToArrayRec(".", "*", $FLTAR_FILES, $FLTAR_NORECUR, $FLTAR_NOSORT, $FLTAR_NOPATH)
	ProgressOn("Postęp", "", "", -1, 0, $DLG_NOTONTOP)
	Local $bEncodedSucessfully = True
	Local $retries = 0
	Local $status = 0
	Local $sleepTime = 50
	Local $GUI = GUICreate("Informacja", 300, 300)
	Local $edit = GUICtrlCreateEdit("", 8, 5, 290, 290, BitOR($ES_AUTOVSCROLL, $ES_READONLY, $WS_VSCROLL))
	ConsoleWriteGUI($edit, "Tych plików nie udało się odkodować - odkoduj ręcznie:" & @CRLF)
	For $i = 1 to UBound($aFileList) - 1
		If StringCompare($aFileList[$i], $decoderName) <> 0 Then
			Do
				$status = LaunchN($aFileList[$i], $sleepTime)
				If $retries = 3 Then
					ConsoleWriteGUI($edit, $aFileList[$i] & @CRLF)
					GUISetState(@SW_SHOW, $GUI)
					$bEncodedSucessfully = False
					ExitLoop
				Endif
				$retries += 1
				$sleepTime += 1000
			Until $status = 1
		Endif
		$retries = 0
		$sleepTime = 50
		ProgressSet(100 * $i / $aFileList[0], "plik " & $i & " / " & ($aFileList[0] - 1))
	Next
	ProgressOff()

	If $bEncodedSucessfully = True Then
		MsgBox($MB_OK, "Dekoder", "Zrobione!")
	Else
		While 1
			Switch GUIGetMsg()
				Case $GUI_EVENT_CLOSE
					ExitLoop
			EndSwitch
		WEnd
	EndIf
EndFunc


Func LaunchN($aFile, $sleepTime)
	Run($notepadPath & " -nosession -alwaysOnTop " & $aFile)
	Local $hWnd = WinWaitActive("[CLASS:Notepad++]", "", 3)
	If $hWnd = 0 Then
		WinKill($hWnd, "")
		Return 0
	Endif
	Sleep($sleepTime)
	Send("!^s")
	Local $sW = WinWaitActive("Zapisywanie jako", "", 3)
	If $sW = 0 Then
		WinKill($hWnd, "")
		Return 0
	Endif
	$status = ControlSetText($sW, "[CLASS:#32770]", "[CLASS:Edit; INSTANCE:1]", $outputDir & "\" & $aFile)
	If $status = 0 Then
		WinKill($hWnd, "")
		Return 0
	Endif
	Send("{ENTER}")
	WinWait($hWnd)
	ControlFocus($hWnd, "Notepad++", "")
	WinClose($hWnd)
	Return 1
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
