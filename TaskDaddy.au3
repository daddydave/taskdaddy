#include-once
#include <Array.au3>
#include "TaskDaddyCore.au3"
#include "TaskDaddyArguments.au3"
#include "TaskDaddyGUI.au3"

AutoItSetOption("MustDeclareVars", 1)

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=TaskDaddy.ico
#AutoIt3Wrapper_UseUpx=N
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Global $InputFile, $FHInputFile, $OOutlook, $StrTask, $LineRead
Global $TaskFields[7]
ParseArguments()
$OOutlook = _OutlookOpen()

If @error Then
	MsgBox(16, "Error " & @error, "Unable to open Microsoft Outlook object")
	Exit (5)
EndIf
If $InputFile = "" Then
	If $IsGUI Then
		$StrTask = LoadGUI($ArgTask, "")
		$TaskFields = ParseTask($StrTask)
		MakeTaskOutlook($OOutlook, $TaskFields)
		If @error > 0 Then
			MsgBox(16, "TaskDaddy Error", "Unable to create Outlook task """ & $StrTask & """")
		EndIf
	Else
		$TaskFields = ParseTask($ArgTask)
		MakeTaskOutlook($OOutlook, $TaskFields)
		If @error > 0 Then
			MsgBox(16, "TaskDaddy Error", "Unable to create Outlook task """ & $ArgTask & """")
		EndIf
	EndIf
Else
	; TODO: Merge string processing
	$FHInputFile = FileOpen($InputFile, 0)
	If $FHInputFile = -1 Then
		MsgBox(16, "Error " & @error, "Unable to open file""" & $InputFile & """")
		Exit (51)
	EndIf
	While 1
		$GuiReturnedSkip = False
		$LineRead = FileReadLine($FHInputFile)
        If @error = -1 Then ExitLoop
        If StringStripWS($LineRead, 3) = "" Then
            ContinueLoop
        EndIf

		If $IsGUI Then
			; combo of GUI mode and file input
            GUIDelete($frmMain)
			$StrTask = LoadGUI($LineRead, $InputFile)
			If $GuiReturnedSkip Then
				ContinueLoop
			EndIf
			If $GuiReturnedCancelAll Then
				ExitLoop
			EndIf
			$TaskFields = ParseTask($StrTask)
			MakeTaskOutlook($OOutlook, $TaskFields)
			If @error > 0 Then
				MsgBox(16, "TaskDaddy Error", "Unable to create Outlook task """ & $StrTask & """")
			EndIf
		Else
			$TaskFields = ParseTask($LineRead)
			MakeTaskOutlook($OOutlook, $TaskFields)
			If @error > 0 Then
				MsgBox(16, "TaskDaddy Error", "Unable to create Outlook task """ & $LineRead & """")
			EndIf
		EndIf
	WEnd
	FileClose($FHInputFile)
	Exit (0)
EndIf
Func LaunchHelp()
	ShellExecute("hh.exe", @ScriptDir & "\TaskDaddy.chm")
EndFunc   ;==>LaunchHelp