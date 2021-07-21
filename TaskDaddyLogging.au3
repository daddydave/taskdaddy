AutoItSetOption("MustDeclareVars", 1)

Local $LogFileName = ""
Local $TaskFileName = ""
Global $DDLogging = False

Func DDLog($Line)

	If Not $DDLogging Then
		Return
	EndIf

	If $LogFileName = "" Then
		$LogFileName = LogFileName()
	EndIf
	Local $FH = FileOpen($LogFileName, 33) ;
	If $FH <> -1 Then
		FileWriteLine($FH, $Line)
		FileClose($FH)
	EndIf
EndFunc   ;==>DDLog

Func LogFileName()
	If $LogFileName == "" Then
		$LogFileName = @ScriptName & "." & @YEAR & @MON & @MDAY & "." & @HOUR & @MIN & @SEC & "." & @MSEC & ".log"
	EndIf
	Return $LogFileName
EndFunc   ;==>LogFileName