#include-once
#include "TaskDaddyOutlook.au3"
#include "TaskDaddyLogging.au3"
#include <Array.au3>

AutoItSetOption("MustDeclareVars", 1)

Local $ImportanceFlag[3] = ["", "?", "!"]

Func MakeTaskOutlook($OOutlook, $TaskFields)
	;_OutlookCreateTask() params are is in this order (Outlook object, subject, body, start date, due date, importance, reminder date, categories)
	If _OutlookCreateTask($OOutlook, $TaskFields[0], $TaskFields[1], $TaskFields[2], $TaskFields[3], $TaskFields[4], $TaskFields[5], $TaskFields[6]) = 0 Then
		SetError(1)
	EndIf
EndFunc   ;==>MakeTaskOutlook


Func ParseTask($TaskStr)
	Local $MyArgs
	$MyArgs = StringSplitRegex($TaskStr, "\s") ;split on white space
	Local $Subject = ""
	Local $GotSubject = False
	Local $Body = ""
	Local $Categories = ""
	Local $DueDate = "" ; = end date
	Local $Importance = $olImportanceNormal
	Local $I
	Local $TaskFields[7]
	;$MyArgs[0] = the number of elements in array, so "For $MyArg in $MyArgs" didn't work
	For $I = 1 To $MyArgs[0]
		Switch StringLeft($MyArgs[$I], 1)
			Case "@"
				If $Subject <> "" Then $GotSubject = True
				If $Categories <> "" Then $Categories = $Categories & ","
				$Categories = $Categories & StringRight($MyArgs[$I], StringLen($MyArgs[$I]) - 1)
			Case "#"
				If $Subject <> "" Then $GotSubject = True
				$DueDate = StringRight($MyArgs[$I], StringLen($MyArgs[$I]) - 1)
			Case ":"
				If $Subject <> "" Then $GotSubject = True
				$Body = $Body & " " & StringRight($MyArgs[$I], StringLen($MyArgs[$I]) - 1)
			Case "?"

				If $GotSubject = False Then
					$Subject = StringRight($MyArgs[$I], StringLen($MyArgs[$I]) - 1)

					$Importance = $olImportanceLow
				Else

					DecideSubjectOrBody($MyArgs[$I], $GotSubject, $Subject, $Body)
				EndIf
			Case "!"

				If $GotSubject = False Then

					$Subject = StringRight($MyArgs[$I], StringLen($MyArgs[$I]) - 1)
					$Importance = $olImportanceHigh
				Else

					DecideSubjectOrBody($MyArgs[$I], $GotSubject, $Subject, $Body)
				EndIf
			Case Else
				DecideSubjectOrBody($MyArgs[$I], $GotSubject, $Subject, $Body)
		EndSwitch
	Next
	$Subject = StringStripWS($Subject, 3)
	$Body = StringStripWS($Body, 3)
	$TaskFields[0] = $Subject
	$TaskFields[1] = $Body
	$TaskFields[2] = ""
	$TaskFields[3] = $DueDate
	$TaskFields[4] = $Importance
	$TaskFields[5] = ""
	$TaskFields[6] = $Categories
	Return $TaskFields
EndFunc   ;==>ParseTask

Func FindTaskOutlook($OOutlook, $TaskFields)
	Local $Verified = False
	Local $FoundTasks = _OutlookGetTasks($OOutlook, $TaskFields[0])
	If @error > 0 Then
		DDLog("FindTaskOutlook:  @error=" & @error & " for """ & $TaskFields[0] & """")
		Return False
	Else
		If $FoundTasks[0][0] <> 1 Then
			DDLog("FindTaskOutlook: " & $FoundTasks[0][0] & " task(s) found for """ & $TaskFields[0] & """")
			Return False

		EndIf
	EndIf


	Select
		Case $TaskFields[0] <> $FoundTasks[1][0]
			DDLog("FindTaskOutlook: Subject mismatch for """ & $TaskFields[0] & _
					""" expecting """ & $TaskFields[0] & """ but finding """ & $FoundTasks[1][0] & """")
		Case $TaskFields[1] <> $FoundTasks[1][10]
			DDLog("FindTaskOutlook: Body mismatch for """ & $TaskFields[0] & _
					""" expecting """ & $TaskFields[1] & """ but finding """ & $FoundTasks[1][10] & """")
;~ 		Case $DateNone <> $FoundTasks[1][2]
;~ 			DDLog("FindTaskOutlook: Due date mismatch for """ & $TaskFields[0] & _
;~ 					""" expecting """ & $TaskFields[3] & """ but finding """ & $FoundTasks[1][2] & """")
		Case $TaskFields[4] <> $FoundTasks[1][4]
			DDLog("FindTaskOutlook: Importance mismatch for """ & $TaskFields[0] & _
					""" expecting """ & $TaskFields[4] & """ but finding """ & $FoundTasks[1][4] & """")
		Case Not $TaskFields[6] == $FoundTasks[1][18]
			DDLog("FindTaskOutlook: Categories mismatch for """ & $TaskFields[0] & _
					""" expecting """ & $TaskFields[6] & """ but finding """ & $FoundTasks[1][18] & """")
		Case Else
			$Verified = True
	EndSelect

	Return $Verified
EndFunc ;== >FindTaskOutlook


Func DecideSubjectOrBody($Arg, $GotSubject, ByRef $Subject, ByRef $Body)

	; Resist temptation to set $GotSubject = True in this function
	; It depends on what follows; if it is a bareword, we don't "got" the whole subject yet
	If StringStripWS($Arg, 3) = "" Then
		Return
	EndIf
	If $GotSubject = True Then
		If $Body <> "" Then
			$Body = $Body & " " & $Arg
		Else
			$Body = $Arg
		EndIf
	Else
		If $Subject <> "" Then
			$Subject = $Subject & " " & $Arg
		Else
			$Subject = $Arg
		EndIf

	EndIf
EndFunc   ;==>DecideSubjectOrBody

Func StringSplitRegex($String, $DelimRegex)
	$String = StringRegExpReplace($String, $DelimRegex, Chr(7))
	Local $Result[10]
	$Result = StringSplit($String, Chr(7))
	Return $Result
EndFunc   ;==>StringSplitRegex