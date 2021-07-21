#include-once
#include <Date.au3>

AutoItSetOption("MustDeclareVars", 1)

; Outlook functions altered from 2010-01-13 version of Outlook.au3 by Wooltown
; Original version http://www.autoitscript.com/forum/index.php?showtopic=89321

Global $olImportanceLow = 0
Global $olImportanceNormal = 1
Global $olImportanceHigh = 2

Const $olTaskItem = 3
Const $olFolderTasks=13

Func _OutlookOpen() 
;; slightly altered
	Local $oOutlook = ObjGet("", "Outlook.Application")
	If @error Or Not IsObj($oOutlook) Then
		$oOutlook = ObjCreate("Outlook.Application")
		If @error Or Not IsObj($oOutlook) Then
			Return SetError(1, 0, 0)
		EndIf
	EndIf
	Return $oOutlook
EndFunc   ;==>_OutlookOpen
;

Func _OutlookCreateTask($oOutlook, $sSubject, $sBody = "", $sStartDate = "", $sDueDate = "", $iImportance = $olImportanceNormal, $sReminderDate = "", $sCategories = "")
;; slightly altered
	Local $iRc = 0
	; Made $oOuError global so it can be referred to from the error handler - for testing purposes only
	Global $oOuError = ObjEvent("AutoIt.Error", "_OutlookError")
	Local $oNote = $oOutlook.CreateItem($olTaskItem)
	$oNote.Subject = $sSubject
	$oNote.Body = $sBody

	If $sStartDate <> "" Then
		$oNote.StartDate = $sStartDate
	EndIf

	If $sDueDate <> "" Then
		$oNote.DueDate = $sDueDate
    EndIf

	If $sReminderDate <> "" Then
		$oNote.ReminderTime = $sReminderDate
        $oNote.ReminderSet = True
    EndIf

	$oNote.Importance = $iImportance

	$oNote.Categories = $sCategories

	$oNote.Save()
	$iRc = @error
	If $iRc = 0 Then
		Return 1
	Else
		Return SetError(9, 0, 0)
	EndIf
EndFunc   ;==>_OutlookCreateTask

Func _OutlookGetTasks($oOutlook, $sSubject = "", $sStartDate = "", $sEndDate = "", $sStatus = "", $sWarningClick = "")
	If $sWarningClick <> "" And	FileExists($sWarningClick) = 0 Then
		Return SetError(2, 0, 0)
	Else
		Run($sWarningClick)
	EndIf
	Local $avTasks[1000][19], $sFilter = "", $oFilteredItems
	Local $oOuError = ObjEvent("AutoIt.Error", "_OutlookError")
	$avTasks[0][0] = 0
	$avTasks[0][1] = 0
	Local $oNamespace = $oOutlook.GetNamespace("MAPI")
	Local $oFolder = $oNamespace.GetDefaultFolder($olFolderTasks)
	Local $oColItems = $oFolder.Items
	$oColItems.Sort("[Start]")
	$oColItems.IncludeRecurrences = True 
	If $sSubject <> "" Then
		$sFilter = '[Subject] = "' & $sSubject & '"'
	EndIf
	If $sStartDate <> "" Then
		If Not _DateIsValid($sStartDate) Then Return SetError(1, 0, 0)
		If $sFilter <> "" Then $sFilter &= ' And '
		$sFilter &= '[Start] >= "' & $sStartDate & '"'
	EndIf	
	If $sEndDate <> "" Then
		If Not _DateIsValid($sEndDate) Then Return SetError(1, 0, 0)
		If $sFilter <> "" Then $sFilter &= ' And '
		$sFilter &= '[Due] <= "' & $sEndDate & '"'
	EndIf	
	If $sStatus <> "" Then
		If $sFilter <> "" Then $sFilter &= ' And '
		$sFilter &= '[Status] = "' & $sStatus & '"'
	EndIf
	If $sFilter = "" Then
		$oFilteredItems = $oColItems
	Else
		$oFilteredItems = $oColItems.Restrict($sFilter)
	EndIf
	For $oItem In $oFilteredItems
 			If $avTasks[0][0] = 999 Then
 				SetError (3)
 				Return $avTasks
 			EndIf
 			$avTasks[0][0] += 1
 			$avTasks[$avTasks[0][0]][0] = $oItem.Subject
 			$avTasks[$avTasks[0][0]][1] = $oItem.StartDate
 			$avTasks[$avTasks[0][0]][2] = $oItem.DueDate
 			$avTasks[$avTasks[0][0]][3] = $oItem.Status
 			$avTasks[$avTasks[0][0]][4] = $oItem.Importance
 			$avTasks[$avTasks[0][0]][5] = $oItem.Complete
 			$avTasks[$avTasks[0][0]][6] = $oItem.PercentComplete
 			If $avTasks[$avTasks[0][0]][6] = 0 Then $avTasks[0][1] += 1
 			$avTasks[$avTasks[0][0]][7] = $oItem.ReminderSet
 			$avTasks[$avTasks[0][0]][8] = $oItem.ReminderTime
 			$avTasks[$avTasks[0][0]][9] = $oItem.Owner
 			$avTasks[$avTasks[0][0]][10] = $oItem.Body
 			$avTasks[$avTasks[0][0]][11] = $oItem.DateCompleted
 			$avTasks[$avTasks[0][0]][12] = $oItem.TotalWork
 			$avTasks[$avTasks[0][0]][13] = $oItem.ActualWork
 			$avTasks[$avTasks[0][0]][14] = $oItem.Mileage
 			$avTasks[$avTasks[0][0]][15] = $oItem.BillingInformation
 			$avTasks[$avTasks[0][0]][16] = $oItem.Companies
 			$avTasks[$avTasks[0][0]][17] = $oItem.Delegator
 			$avTasks[$avTasks[0][0]][18] = $oItem.Categories
	Next
	$oItem = ""
	$oColItems = ""
	$oFolder = ""
	$oNamespace = ""
	If $avTasks[0][0] = 0 Then Return SetError(2, 0, 0)
	Redim $avTasks[$avTasks[0][0] + 1][19]
	Return $avTasks
EndFunc