#include-once
#include <Array.au3>
#include <_OptParse.au3>
#include "TaskDaddyCore.au3"
#include "TaskDaddyArgs.au3"
#include "TaskDaddyGUI.au3"

AutoItSetOption("MustDeclareVars", 1)

; TaskDaddy
; a NANY 2011 project (http://www.donationcoder.com/Forums/bb/index.php?topic=22721.0)
; Copyright 2010 by David Eason (daddydave)
; Except for the following useful functions:
; _OptParse.au3 by Stephen Podhajecki (eltorro)
; http://code.google.com/p/my-autoit/downloads/detail?name=_OptParse_031209.zip
; Outlook.au3 by Wooltown (altered version)
; http://www.autoitscript.com/forum/index.php?showtopic=89321
; Inspired by Bob Menke's Add Task VBScript:
; http://www.outlookcode.com/codedetail.aspx?id=1642
; Uses an altered version of Outlook.au3 (2010-01-13) written by Wooltown
; The following alterations made (altered code follows):
;~ Func _OutlookOpen()
;~ 	Local $oOutlook = ObjGet("", "Outlook.Application")
;~ 	If @error Or Not IsObj($oOutlook) Then
;~ 		$oOutlook = ObjCreate("Outlook.Application")
;~ 		If @error Or Not IsObj($oOutlook) Then
;~ 			Return SetError(1, 0, 0)
;~ 		EndIf
;~ 	EndIf
;~ 	Return $oOutlook
;~ EndFunc   ;==>_OutlookOpen
;~ Func _OutlookCreateTask($oOutlook, $sSubject, $sBody = "", $sStartDate = "", $sDueDate = "", $iImportance = $olImportanceNormal, $sReminderDate = "", $sCategories = "")
;~ 	Local $iRc = 0
;~ 	; Made $oOuError global so it can be referred to from the error handler - for testing purposes only
;~ 	Global $oOuError = ObjEvent("AutoIt.Error", "_OutlookError")
;~ 	Local $oNote = $oOutlook.CreateItem($olTaskItem)
;~ 	$oNote.Subject = $sSubject
;~ 	$oNote.Body = $sBody
;~ 	If $sStartDate <> "" Then
;~ 		$oNote.StartDate = $sStartDate
;~ 	EndIf
;~ 	If $sDueDate <> "" Then
;~ 		$oNote.DueDate = $sDueDate
;~     EndIf
;~ 	If $sReminderDate <> "" Then
;~ 		$oNote.ReminderTime = $sReminderDate
;~         $oNote.ReminderSet = True
;~     EndIf
;~ 	$oNote.Importance = $iImportance
;~ 	$oNote.Categories = $sCategories
;~ 	$oNote.Save()
;~ 	$iRc = @error
;~ 	If $iRc = 0 Then
;~ 		Return 1
;~ 	Else
;~ 		Return SetError(9, 0, 0)
;~ 	EndIf
;~ EndFunc   ;==>_OutlookCreateTask
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=TaskDaddy.ico
#AutoIt3Wrapper_UseUpx=N
#AutoIt3Wrapper_Run_After=copy TaskDaddy.exe c:\tools\TaskDaddy
#AutoIt3Wrapper_Run_After=copy help\TaskDaddy.chm c:\tools\TaskDaddy
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

Global $InputFile, $FHInputFile, $OOutlook, $StrTask, $LineRead
Global $TaskFields[7]
ParseArgs()
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