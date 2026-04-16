#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <ScreenCapture.au3>
#include <GDIPlus.au3>
#include <ColorConstants.au3>
#include "UWPOCR.au3"
#include <WinAPISysWin.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <StaticConstants.au3>
#include <File.au3>

;#Include <Misc.au3>
#AutoIt3Wrapper_Res_HiDpi=Y
opt("PixelCoordMode", 1)
Global $g_tStruct = DllStructCreate($tagPOINT) ; Create a structure that defines the point to be checked.
Global $aselxpos[0], $aselypos[0], $aselwidth[0], $aselheight[0], $aseldesc[0],$List2Sel=0
If Not (@Compiled ) Then DllCall("User32.dll","bool","SetProcessDPIAware")

Local $sinifile=@ScriptDir&'\'& StringRegExpReplace(@ScriptName,'.([^.]*$)','.ini'),$xpos, $ypos, $winwidth,$winheight, $selxpos, $yselpos, $selwinwidth,$selwinheight, $hWnd
$xpos = IniRead($sinifile, "Settings", "xpos","100")
$ypos = IniRead($sinifile, "Settings", "ypos","100")
$winwidth = IniRead($sinifile, "Settings", "winwidth","578")
$winheight = IniRead($sinifile, "Settings", "winheight","189")
$selxpos = IniRead($sinifile, "Settings", "selxpos","100")
$selypos = IniRead($sinifile, "Settings", "selypos","100")
$selwinwidth = IniRead($sinifile, "Settings", "selwinwidth","578")
$selwinheight = IniRead($sinifile, "Settings", "selwinheight","189")
$ReceivingWindow=0

Local $hFileOpen = FileOpen($sinifile, $FO_READ)
$numberofpresets=_FileCountLines ($hFileOpen)-10;
FileClose($hFileOpen)

ReadPresets()

$numberofpresets=UBound($aselxpos)

$hWnd = GUICreate("OCR Typer", 377, 239, 251, 224,BitOr($WS_MINIMIZEBOX, $WS_SIZEBOX), $WS_EX_TOPMOST)
$List1 = GUICtrlCreateList("", 8, 8, 41, 188, BitOR($LBS_NOTIFY,$WS_VSCROLL,$WS_BORDER))
GUICtrlSetData(-1, ".5X|1X|2X|3X|4X|5X|6X|7X|8X|9X", BitOR($WS_BORDER, $WS_VSCROLL))
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$List2 = GUICtrlCreateList("", 56, 8, 129, 188, BitOR($LBS_NOTIFY,$WS_VSCROLL,$WS_BORDER))
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
GUICtrlSetData(-1, "")
$Button1 = GUICtrlCreateButton("ON", 264, 8, 105, 170)
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button2 = GUICtrlCreateButton("+", 192, 8, 33, 57)
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button3 = GUICtrlCreateButton("-", 224, 8, 33, 57)
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button4 = GUICtrlCreateButton("↓", 224, 64, 33, 57)
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button5 = GUICtrlCreateButton("↑", 192, 64, 33, 57)
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button6 = GUICtrlCreateButton("REN", 192, 120, 33, 57)
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button7 = GUICtrlCreateButton("SAV", 224, 120, 33, 57)
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
GUISetState(@SW_SHOW)
$clickdisabled=0
WinMove($hWnd,  '',$xpos, $ypos, $winwidth,$winheight)

For $i=0 to $numberofpresets-1
	_GUICtrlListBox_AddString($List2,  $aseldesc[$i])
	_GUICtrlListBox_SetTopIndex($List2, $i)
	;_GUICtrlListBox_ClickItem($List2, $i, "left")
Next

Local $hRedBox = GUICreate("", 100, 100, -1, -1, BitOr($WS_POPUP, $WS_SIZEBOX), $WS_EX_TOPMOST)
GUISetBkColor(0xFF0000)
WinSetTrans($hRedBox, "", 20)
Local $iLabel = GUICtrlCreateLabel("", 0, 0, 100, 100, -1, $GUI_WS_EX_PARENTDRAG)
GUISetState(@SW_SHOW)
GUIRegisterMsg($WM_EXITSIZEMOVE, "on_WM_EXITSIZEMOVE")
WinMove($hRedBox, "", $selxpos, $selypos, $selwinwidth, $selwinheight)

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
	Case $GUI_EVENT_CLOSE
			GUIDelete()
			;DllClose($hDLL)
			OnAutoItExit()
			Exit
		Case $GUI_EVENT_PRIMARYDOWN
			$clickdisabled=1
        Case $iLabel
			if Not WinExists($ReceivingWindow) Then 
				$clickdisabled=1
				GUICtrlSetStyle($Button1, $GUI_SS_DEFAULT_BUTTON)
				ContinueLoop
			EndIf
			if Not $clickdisabled Then
				Local $aClientSize = WinGetPos ($hRedBox); 
				GUISetState(@SW_HIDE,$hRedBox)			
				Local $text=OCRBitmap($aClientSize[0],$aClientSize[1],$aClientSize[0]+$aClientSize[2],$aClientSize[1]+$aClientSize[3]) 
				$text=StringReplace($text,@CRLF,"")
				GUISetState(@SW_SHOW,$hRedBox)
				WinActivate($ReceivingWindow)
				SendKeepActive($ReceivingWindow,'')
				Send($text)
			EndIf
		Case $GUI_EVENT_PRIMARYUP
			$clickdisabled=0
        Case $List1   
            Local $List1Sel=0
            local $index=0
            $List1Sel=_GUICtrlListBox_GetCurSel($List1)
			Local $aClientSize = WinGetPos ($hRedBox);WinGetClientSize($hRedBox)
			Switch $List1Sel
				Case 0
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(($aselwidth[$List2Sel]*1)/2), ($aClientSize[1]+$aClientSize[3]/2)-(($aselheight[$List2Sel]*.5)/2), $aselwidth[$List2Sel]*.5,$aselheight[$List2Sel]*.5)
				Case 1
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(($aselwidth[$List2Sel]*1)/2), ($aClientSize[1]+$aClientSize[3]/2)-(($aselheight[$List2Sel]*1)/2), $aselwidth[$List2Sel]*1,$aselheight[$List2Sel]*1)
				Case 2
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(($aselwidth[$List2Sel]*2)/2), ($aClientSize[1]+$aClientSize[3]/2)-(($aselheight[$List2Sel]*2)/2), $aselwidth[$List2Sel]*2,$aselheight[$List2Sel]*2)
				Case 3
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(($aselwidth[$List2Sel]*3)/2), ($aClientSize[1]+$aClientSize[3]/2)-(($aselheight[$List2Sel]*3)/2), $aselwidth[$List2Sel]*3,$aselheight[$List2Sel]*3)
				Case 4
					WinMove($hRedBox, "",  ($aClientSize[0]+$aClientSize[2]/2)-(($aselwidth[$List2Sel]*4)/2), ($aClientSize[1]+$aClientSize[3]/2)-(($aselheight[$List2Sel]*4)/2), $aselwidth[$List2Sel]*4,$aselheight[$List2Sel]*4)
				Case 5
					WinMove($hRedBox, "",  ($aClientSize[0]+$aClientSize[2]/2)-(($aselwidth[$List2Sel]*5)/2), ($aClientSize[1]+$aClientSize[3]/2)-(($aselheight[$List2Sel]*5)/2), $aselwidth[$List2Sel]*5,$aselheight[$List2Sel]*5)
				Case 6
					WinMove($hRedBox, "",  ($aClientSize[0]+$aClientSize[2]/2)-(($aselwidth[$List2Sel]*6)/2), ($aClientSize[1]+$aClientSize[3]/2)-(($aselheight[$List2Sel]*6)/2), $aselwidth[$List2Sel]*6,$aselheight[$List2Sel]*6)	
				Case 7
					WinMove($hRedBox, "",  ($aClientSize[0]+$aClientSize[2]/2)-(($aselwidth[$List2Sel]*7)/2), ($aClientSize[1]+$aClientSize[3]/2)-(($aselheight[$List2Sel]*7)/2), $aselwidth[$List2Sel]*7,$aselheight[$List2Sel]*7)
				Case 8
					WinMove($hRedBox, "",  ($aClientSize[0]+$aClientSize[2]/2)-(($aselwidth[$List2Sel]*8)/2), ($aClientSize[1]+$aClientSize[3]/2)-(($aselheight[$List2Sel]*8)/2), $aselwidth[$List2Sel]*8,$aselheight[$List2Sel]*8)
				Case 9
					WinMove($hRedBox, "",  ($aClientSize[0]+$aClientSize[2]/2)-(($aselwidth[$List2Sel]*9)/2), ($aClientSize[1]+$aClientSize[3]/2)-(($aselheight[$List2Sel]*9)/2), $aselwidth[$List2Sel]*9,$aselheight[$List2Sel]*9)
			EndSwitch
        Case $List2  
            local $index=0
			CaptureClickToggle(1)
			_GUICtrlListBox_ClickItem($List1, 1, "left")
            $List2Sel=_GUICtrlListBox_GetCurSel($List2)
			if @error Then $List2Sel=0
				ConsoleWrite('selected: '&$List2Sel&' size: '&$aselwidth[$List2Sel]&' '&$aselheight[$List2Sel]&@CRLF)
				;_ArrayDisplay($aselheight)
			WinMove($hRedBox, "", $aselxpos[$List2Sel], $aselypos[$List2Sel], $aselwidth[$List2Sel], $aselheight[$List2Sel])
			CaptureClickToggle(0)
	Case $Button1
		GUICtrlSetBkColor($Button1, $COLOR_ORANGE)
		While 1
			Sleep(10)
			$ReceivingWindow = PointGetHandle()
			if  $ReceivingWindow  = $hRedBox OR $ReceivingWindow  = $hWnd Or StringInStr(WinGetTitle( $ReceivingWindow,""),'au3') Then
			Else
				$clickdisabled=0
				GUICtrlSetBkColor($Button1, $COLOR_GREEN)
				;_WinAPI_GetProcessNameFromHWND($hWND)
				;ConsoleWrite(WinGetTitle( $ReceivingWindow,"")&@CRLF)
				ExitLoop
			EndIf
		WEnd
	Case $Button2 ;ADD PRESET
		Local $preset = "475|117|450|67|Preset "
		Local $selecteditem=_GUICtrlListBox_GetCaretIndex($List2),$numitems=_GUICtrlListBox_GetCount($List2)
		CaptureClickToggle(1)
		$preset=StringSplit($preset,'|')
		$preset[5]=InputBox("Enter Preset Name", "Please input the item name ", $preset[5]&UBound($aselxpos)+1)
		if @error Then ContinueLoop
		 _GUICtrlListBox_AddString($List2,  $preset[5])
		_GUICtrlListBox_SetTopIndex($List2, UBound($aselxpos))
		_GUICtrlListBox_ClickItem($List2, UBound($aselxpos), "left")
		_GUICtrlListBox_ClickItem($List1, 1, "left")
		_ArrayAdd($aselxpos,$preset[1])
		_ArrayAdd($aselypos,$preset[2])
		_ArrayAdd($aselwidth,$preset[3])
		_ArrayAdd($aselheight,$preset[4])
		_ArrayAdd($aseldesc,$preset[5])
		CaptureClickToggle(1)
	Case $Button3
		Local $selecteditem=_GUICtrlListBox_GetCaretIndex($List2),$numitems=_GUICtrlListBox_GetCount($List2)
		If $numitems<=0 Then ContinueLoop
		CaptureClickToggle(1)
			_GUICtrlListBox_DeleteString($List2, $selecteditem)	
			_ArrayDelete($aselxpos, $selecteditem)
			_ArrayDelete($aselypos, $selecteditem)
			_ArrayDelete($aselwidth, $selecteditem)
			_ArrayDelete($aselheight, $selecteditem)
			_ArrayDelete($aseldesc, $selecteditem)
			$numberofpresets=UBound($aselxpos)
			if $selecteditem>=$numitems-1 then 
			$selecteditem=$numitems-2
		EndIf
		_GUICtrlListBox_SetTopIndex($List2, $selecteditem)
		_GUICtrlListBox_ClickItem($List2, $selecteditem, "left")
		CaptureClickToggle(0)
	Case $Button4
		Local $selecteditem=_GUICtrlListBox_GetCurSel($List2),$newtext
		If $selecteditem>=_GUICtrlListBox_GetCount($List2)-1 Then ContinueLoop
		CaptureClickToggle(1)
		Local $text1=_GuiCtrlListbox_GetText($List2, $selecteditem)
		Local $text2=_GuiCtrlListbox_GetText($List2, $selecteditem+1)
		_GUICtrlListBox_ReplaceString($List2, $selecteditem, $text2)
		_GUICtrlListBox_ReplaceString($List2, $selecteditem+1, $text1)
		_ArraySwap($aselxpos, $selecteditem,$selecteditem+1)
		_ArraySwap($aselypos, $selecteditem,$selecteditem+1)
		_ArraySwap($aselwidth, $selecteditem,$selecteditem+1)
		_ArraySwap($aselheight, $selecteditem,$selecteditem+1)
		_ArraySwap($aseldesc, $selecteditem,$selecteditem+1)
		_GUICtrlListBox_SetSel($List2,$selecteditem,0)
		_GUICtrlListBox_SetTopIndex($List2, $selecteditem+1)
		_GUICtrlListBox_ClickItem($List2, $selecteditem+1, "left")
		CaptureClickToggle(0)
	Case $Button5
		Local $selecteditem=_GUICtrlListBox_GetCurSel($List2),$newtext
		If $selecteditem<1 Then ContinueLoop
		CaptureClickToggle(1)
		Local $text1=_GuiCtrlListbox_GetText($List2, $selecteditem)
		Local $text2=_GuiCtrlListbox_GetText($List2, $selecteditem-1)
		_GUICtrlListBox_ReplaceString($List2, $selecteditem, $text2)
		_GUICtrlListBox_ReplaceString($List2, $selecteditem-1, $text1)
		_ArraySwap($aselxpos, $selecteditem,$selecteditem-1)
		_ArraySwap($aselypos, $selecteditem,$selecteditem-1)
		_ArraySwap($aselwidth, $selecteditem,$selecteditem-1)
		_ArraySwap($aselheight, $selecteditem,$selecteditem-1)
		_ArraySwap($aseldesc, $selecteditem,$selecteditem-1)
		_GUICtrlListBox_SetSel($List2,$selecteditem,0)
		_GUICtrlListBox_SetTopIndex($List2, $selecteditem)
		_GUICtrlListBox_ClickItem($List2, $selecteditem-1, "left")
		CaptureClickToggle(0)
	Case $Button6
		Local $selecteditem=_GUICtrlListBox_GetCurSel($List2),$newtext
		If $selecteditem=-1 Then ContinueLoop
		CaptureClickToggle(1)
		$aseldesc[$selecteditem]=InputBox("Rename Preset", "Please inout the new name for item "&$selecteditem, _GuiCtrlListbox_GetText($List2, $selecteditem))
		If StringLen($aseldesc[$selecteditem])<1 Then ContinueCase
		_GUICtrlListBox_ReplaceString($List2, $selecteditem, $aseldesc[$selecteditem])
		_GUICtrlListBox_SetTopIndex($List2, $selecteditem)
		_GUICtrlListBox_ClickItem($List2, $selecteditem, "left")
		CaptureClickToggle(0)
	Case $Button7
		SavePresets()
	EndSwitch
WEnd

Func OCRBitmap($x,$y,$height,$width)
	_GDIPlus_Startup()
	Local $hTimer = TimerInit()
	Local $hHBitmap = _ScreenCapture_Capture("", $x,$y,$height,$width, False)
	Local $hBitmap = _GDIPlus_BitmapCreateFromHBITMAP($hHBitmap)
	Local $sOCRTextResult = _UWPOCR_GetText($hBitmap, Default, True)
	ConsoleWrite($sOCRTextResult&@CRLF)
	_WinAPI_DeleteObject($hHBitmap)
	_GDIPlus_BitmapDispose($hBitmap)
	_GDIPlus_Shutdown()
	Return $sOCRTextResult
EndFunc

func OnAutoItExit()
    IniWrite($sinifile, "Settings", "xpos", $xpos)
    IniWrite($sinifile, "Settings", "ypos", $ypos)
	IniWrite($sinifile, "Settings", "winwidth",$winwidth)
    IniWrite($sinifile, "Settings", "winheight",$winheight)
    IniWrite($sinifile, "Settings", "selxpos", $selxpos)
    IniWrite($sinifile, "Settings", "selypos", $selypos)
	IniWrite($sinifile, "Settings", "selwinwidth",$selwinwidth)
    IniWrite($sinifile, "Settings", "selwinheight",$selwinheight)	
    ;if $sComPort<>$oldport Then IniWrite($sinifile, "Settings", "port", $sComPort)
EndFunc

Func on_WM_EXITSIZEMOVE($_hWnd, $msg, $wParam, $lParam)
    If $_hWnd = $hWnd Then
		sleep(10)
       $a = WinGetPos($hWnd)
		$xpos=$a[0]
		$ypos=$a[1]
		$winwidth=$a[2]
		$winheight=$a[3]
		ElseIf $_hWnd = $hRedBox Then
		sleep(10)
       $a = WinGetPos($hRedBox)
		$selxpos=$a[0]
		$selypos=$a[1]
		$selwinwidth=$a[2]
		$selwinheight=$a[3]
		$aselxpos[$List2Sel]=$a[0]
		$aselypos[$List2Sel]=$a[1]
		$aselwidth[$List2Sel]=$a[2]
		$aselheight[$List2Sel]=$a[3]
		;ConsoleWrite('moved, size: '&$aselwidth[$List2Sel]&' '&$aselheight[$List2Sel]&@CRLF)
	EndIf
EndFunc
Func CaptureClickToggle($toggle)
	$clickdisabled=$toggle
	switch $toggle
	Case 1
		GUISetState(@SW_HIDE,$hRedBox)
	Case Else
		GUISetState(@SW_SHOW,$hRedBox)
	EndSwitch
EndFunc 

Func ReadPreset($num)
$preset = IniRead($sinifile, "Settings", "p"&$num,"475|117|450|67|Preset 1")
$preset=StringSplit($preset,'|')
_ArrayAdd($aselxpos,$preset[1])
_ArrayAdd($aselypos,$preset[2])
_ArrayAdd($aselwidth,$preset[3])
_ArrayAdd($aselheight,$preset[4])
_ArrayAdd($aseldesc,$preset[5])
;ConsoleWrite($aseldesc[$num]&@CRLF)
;	Return $hWindow
EndFunc 

Func ReadPresets()
	For $i=0 to $numberofpresets
		$preset = IniRead($sinifile, "Settings", "p"&$i,"475|117|450|67|Preset 1")
		$preset=StringSplit($preset,'|')
		_ArrayAdd($aselxpos,$preset[1])
		_ArrayAdd($aselypos,$preset[2])
		_ArrayAdd($aselwidth,$preset[3])
		_ArrayAdd($aselheight,$preset[4])
		_ArrayAdd($aseldesc,$preset[5])
	Next	
EndFunc 

Func SavePresets()
	Local $inifilebuffer=''
	Local $hFileOpen = FileOpen($sinifile, $FO_READ)
    If $hFileOpen = -1 Then
		$inifilebuffer=StringReplace('[Settings]|xpos=890|ypos=1158|winwidth=719|winheight=33|selxpos=1516|selypos=547|selwinwidth=264|selwinheight=72','|',@CRLF)
	Else
		For $i=1 to 10
			If StringLen($inifilebuffer)>1 Then
				$inifilebuffer=$inifilebuffer&@CRLF&FileReadLine($hFileOpen)
			Else
				$inifilebuffer=FileReadLine($hFileOpen)
			EndIf
		Next
	FileClose($hFileOpen)		
	EndIf

	$hFileOpen = FileOpen($sinifile, $FO_OVERWRITE)
	FileWrite($hFileOpen,$inifilebuffer)
	FileClose($hFileOpen)
	
	For $i=0 to UBound($aselxpos)-1
		IniWrite($sinifile, "Settings", "p"&$i,$aselxpos[$i]&'|'&$aselypos[$i]&'|'&$aselwidth[$i]&'|'&$aselheight[$i]&'|'&$aseldesc[$i])
	Next	
EndFunc 

Func PointGetHandle()
	Local $hWindow=0
	Local $WindowTitle=''
        While 1
				DllStructSetData($g_tStruct, "x", MouseGetPos(0))
				DllStructSetData($g_tStruct, "y", MouseGetPos(1))
                $hWindow = _WinAPI_WindowFromPoint($g_tStruct) 
				$hWindow=_WinAPI_GetAncestor ( $hWindow,$GA_ROOTOWNER )
				$WindowTitle=WinGetTitle ( $hWindow,'' )
               ; ToolTip(WinGetTitle ( $hWindow,'' )&' '&$hWindow) 
				if StringLen($WindowTitle)>0 Then 
					ExitLoop
				EndIf
                Sleep(50)
		WEnd
	Return $hWindow
EndFunc 