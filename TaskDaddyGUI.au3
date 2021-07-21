; TaskDaddy
; a NANY 2011 project (http://www.donationcoder.com/Forums/bb/index.php?topic=22721.0)
; Copyrighted 2010 by David Eason
; Except for the following useful functions:
; _OptParse.au3 by Stephen Podhajecki (eltorro)
; http://code.google.com/p/my-autoit/downloads/detail?name=_OptParse_031209.zip
; Outlook.au3 by Wooltown (altered version)
; http://www.autoitscript.com/forum/index.php?showtopic=89321
; Uses an altered version of Outlook.au3 (2010-01-13) written by Wooltown
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
; Used for file input GUI mode (/f /p)
Global $GuiReturnedSkip, $GuiReturnedCancelAll
Global $frmMain
$GuiReturnedSkip = False
$GuiReturnedCancelAll = False
Func LoadGUI($StrTask = "", $InputFile = "")
	If $InputFile = "" Then
		$frmMain = GUICreate("Enter task to create - TaskDaddy", 583, 99, 557, 650)
	Else
		$frmMain = GUICreate("Confirm or modify task to create - TaskDaddy", 583, 99, 557, 650)
	EndIf
	#Region ### START Koda GUI section ### Form=C:\keep\private\code\au3\TaskDaddy\MainForm.kxf
	GUISetIcon("", -1)
	GUISetFont(9, 400, 0, "")
	Local $Input1 = GUICtrlCreateInput("", 16, 32, 497, 23, $GUI_SS_DEFAULT_INPUT)
	GUICtrlSetFont(-1, 9, 400, 0, "Arial")
	Local $Button1 = GUICtrlCreateButton("", 528, 40, 16, 16, BitOR($BS_DEFPUSHBUTTON, $BS_CENTER, $BS_NOTIFY, $BS_ICON))
	GUICtrlSetImage(-1, @SystemDir & "\shell32.dll", -177, 0)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	Local $Button2 = GUICtrlCreateButton("", 552, 40, 16, 16, $BS_ICON)
	GUICtrlSetImage(-1, @SystemDir & "\shell32.dll", -24, 0)
	Local $Label1 = GUICtrlCreateLabel("To create an Outlook task now, type it below and press Enter. Press F1 for help.", 17, 8, 436, 19)
	Local $frmMain_AccelTable[1][2] = [["{F1}", $Button2]]
	#EndRegion ### END Koda GUI section ###
	;TODO: Change form path
	Local $Button3, $Button4
	If $InputFile <> "" Then
		$Button3 = GUICtrlCreateButton("S&kip", 272, 64, 115, 25)
		$Button4 = GUICtrlCreateButton("&Cancel all", 400, 64, 115, 25)
	EndIf
	GUISetAccelerators($frmMain_AccelTable)
	GUISetState(@SW_SHOW)
	GUICtrlSetData($Input1, $StrTask)
	Local $nMsg
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $Input1
			Case $Button1
				$StrTask = GUICtrlRead($Input1)
				Return $StrTask
			Case $Button2
				ShellExecute("hh.exe", @ScriptDir & "\TaskDaddy.chm")
			Case $Button3
				If $InputFile <> "" Then
					$GuiReturnedSkip = True
					Return ""
				EndIf
			Case $Button4
				If $InputFile <> "" Then
					$GuiReturnedCancelAll = True
					Return ""
				EndIf
		EndSwitch
	WEnd
EndFunc   ;==>LoadGUI