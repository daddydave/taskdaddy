#include-once
#include <Array.au3>
AutoItSetOption("MustDeclareVars", 1)

Global $inputFile = ""
Global $isGUI = False
Global $argTask

Func ShowTaskDaddyUsage($isError = False)
	
	Local $usage =    "Adds task(s) to add to Outlook, either from command line, GUI prompt, or a file." & @CRLF & @CRLF & _
				"taskdaddy [task] [/f [drive:][path]filename] [/p] [/?]" & @CRLF & @CRLF & _
				"/f [drive:][path]filename Specify file containing tasks, one per line" & @CRLF & _
				"/p Force GUI prompt (default)" & @CRLF & _
				"/? Show usage" & @CRLF & @CRLF & _
				"Examples:" & @CRLF & _
				"taskdaddy @Yard Plant murraya tree  " & @CRLF & _
				"Creates single Outlook task with category Yard" & @CRLF & @CRLF & _
				"taskdaddy /f checklist.txt" & @CRLF & _      
				"Creates tasks from a file, one task per line" & @CRLF & @CRLF & _
				"taskdaddy or taskdaddy /p" & @CRLF & _       
				"Prompts for task to enter" & @CRLF & @CRLF & _
				"taskdaddy /p @Taxes Meet with" & @CRLF & _     
				"Prompts with task partially prefilled"
	
	MsgBox(16, "TaskDaddy Usage", $usage)
	If ($isError) Then
		Exit(98)
	Else
		Exit(0)
	EndIf
EndFunc    ;==>ShowTaskDaddyUsage

Func ParseArguments()
	
	Local $ndx, $exceptionCode

	If $CmdLine[0] = 0 Then
		Exit(0)
	EndIf

	$argTask = ""

	; Merge string was never implemented previously, /m option processing removed

	; Error handling technique: https://www.autoitscript.com/forum/topic/70669-exception-handling-in-autoit/
	For $ndx = 1 To $CmdLine[0]

		Switch ($CmdLine[$ndx])
			Case "/f", "-f"
				;;; TRY
				Do 
					$inputFile = $CmdLine[$ndx + 1]
					$exceptionCode = 23
					ExitLoop
				Until 1

				;;; CATCH
				Switch $exceptionCode
					Case 0      ; no exception
					Case 23
						MsgBox(16, "TaskDaddy Error", "/f must specify file containing tasks)")
						Exit(23)
				EndSwitch

			Case "/p", "-p"
				$isGUI = True
			Case "/?", "-?"
				ShowTaskDaddyUsage()
			Case Else
				If $argTask = "" Then 
					$argTask = $CmdLine[$ndx]
				Else 
					$argTask = $argTask & " " & $CmdLine[$ndx]
				EndIf
		EndSwitch
	Next

	If ($inputFile <> "" And $argTask <> "") Then
		MsgBox(16, "TaskDaddy Error", "If using task input file (/f), do not specify task on command line")
		Exit(2)
	EndIf

	If ($inputFile = "" And $argTask = "") Then
		$isGUI = True
	EndIf




EndFunc    ;==>ParseArguments
