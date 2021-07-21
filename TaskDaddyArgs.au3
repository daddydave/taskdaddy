; TaskDaddy
; a NANY 2011 project (http://www.donationcoder.com/Forums/bb/index.php?topic=22721.0)
; Copyrighted 2010 by David Eason
; Except for the following useful functions:
; _OptParse.au3 by Stephen Podhajecki (eltorro)
; http://code.google.com/p/my-autoit/downloads/detail?name=_OptParse_031209.zip
; Outlook.au3 by Wooltown (altered version)
; http://www.autoitscript.com/forum/index.php?showtopic=89321

; Uses an altered version of Outlook.au3 (2010-01-13) written by Wooltown

#include "_OptParse.au3"
#include "Array.au3"
AutoItSetOption("MustDeclareVars", 1)
Global $InputFile, $MergeString, $ArgTask, $IsGUI

Func ParseArgs()
	Global $ValidOpts[10], $Opts[10]
	Local $Ndx, $I
	Local $CmdLn = $CmdLine
	;initialize options parser
	_OptParse_Init($ValidOpts, _
			"TaskDaddy\n", _
			"Copyright(c) 2010, David Eason\n", _
			"taskdaddy ""task specification"" [options]\n\n" & _
			"Leave off options or use /p to create task from GUI\n\n")
	_OptParse_Add($ValidOpts, "f", "", $OPT_ARG_REQ, "Input file containing a list of tasks")
	_OptParse_Add($ValidOpts, "m", "", $OPT_ARG_REQ, "Optional merge string to be used with input file") ;TODO
	_OptParse_Add($ValidOpts, "p", "", $OPT_ARG_NONE, "Prompt user before creating each task")
	_OptParse_Add($ValidOpts, "?", "help", $OPT_ARG_NONE, "Display command line options")
	_OptParse_SetDisplay(1) ; 0= console, 1= msgbox
	$Opts = _OptParse_GetOpts($CmdLn, $ValidOpts)

	If @error > 1 Then
		_OptParse_ShowUsage($ValidOpts, 1)
		Exit
	EndIf
	If _OptParse_MatchOption("?", $Opts, $Ndx) Then
		_OptParse_ShowUsage($ValidOpts, 1)
		Exit (98)
	EndIf
	If _OptParse_MatchOption("f", $Opts, $Ndx) Then
		$InputFile = $Opts[$Ndx][1]
	Else
		$InputFile = ""
	EndIf
	If _OptParse_MatchOption("p", $Opts, $Ndx) Then
		$IsGUI = True
	EndIf
	If _OptParse_MatchOption("m", $Opts, $Ndx) Then
		If $InputFile <> "" Then
			$MergeString = $Opts[$Ndx][1]
		Else
			MsgBox(16, "TaskDaddy Error", "Unable to use merge string (/m) without input file (/f)")
			Exit (1)
		EndIf
	Else
		$MergeString = ""
	EndIf
	Select
		Case $CmdLn[0] = 0
			$ArgTask = ""
		Case $CmdLn[0] = 1
			If $InputFile <> "" Then
				MsgBox(16, "TaskDaddy Error", "Command line task not allowed with input file (/f)")
				Exit (2)
			Else
				$ArgTask = $CmdLn[1]
			EndIf
		Case $CmdLn[0] > 1
			For $I = 1 To $CmdLn[0]
				$ArgTask = $ArgTask & " " & $CmdLn[$I]
			Next
		Case Else
	EndSelect
	If ($InputFile = "" And $ArgTask = "") Then
		$IsGUI = True
	EndIf
EndFunc   ;==>ParseArgs