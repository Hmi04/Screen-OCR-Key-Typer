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

;#Include <Misc.au3>
#AutoIt3Wrapper_Res_HiDpi=Y
opt("PixelCoordMode", 1)
Global $g_tStruct = DllStructCreate($tagPOINT) ; Create a structure that defines the point to be checked.

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

$hWnd = GUICreate("OCR Key Typer", 520, 196, 0, 0,BitOr($WS_MINIMIZEBOX, $WS_SIZEBOX), $WS_EX_TOPMOST)
$List1 = GUICtrlCreateList("", 8, 0, 113, 158, BitOR($LBS_NOTIFY,$WS_VSCROLL,$WS_BORDER))
GUICtrlSetData(-1, "1X|2X|3X|4X|5X|6X|7X|8X|9X", BitOR($WS_BORDER, $WS_VSCROLL))
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$List2 = GUICtrlCreateList("", 128, 0, 273, 158, BitOR($LBS_NOTIFY,$WS_VSCROLL,$WS_BORDER))
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
GUICtrlSetData(-1, "Square|Rectangle|Double Wide|Triple Wide|Quad Wide|Pentuple|Hextuple")
$Button1 = GUICtrlCreateButton("ON", 408, 0, 105, 145)
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
GUISetState(@SW_SHOW)
$clickdisabled=0
WinMove($hWnd,  '',$xpos, $ypos, $winwidth,$winheight)

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
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(100/2),($aClientSize[1]+$aClientSize[3]/2)-(100/2), 100,100)
				Case 1
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(200/2),($aClientSize[1]+$aClientSize[3]/2)-(200/2), 200,200 )
				Case 2
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(300/2), ($aClientSize[1]+$aClientSize[3]/2)-(300/2), 300 ,300)
				Case 3
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(400/2), ($aClientSize[1]+$aClientSize[3]/2)-(400/2), 400 ,400)
				Case 4
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(500/2),($aClientSize[1]+$aClientSize[3]/2)-(500/2),500 ,500)
				Case 5
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(600/2), ($aClientSize[1]+$aClientSize[3]/2)-(600/2), 600,600)			
				Case 6
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(700/2), ($aClientSize[1]+$aClientSize[3]/2)-(700/2), 700,700)
				Case 7
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(800/2), ($aClientSize[1]+$aClientSize[3]/2)-(800/2), 800 ,800)	
				Case 8
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(900/2), ($aClientSize[1]+$aClientSize[3]/2)-(900/2), 900 ,900)	
				Case 9
					WinMove($hRedBox, "", ($aClientSize[0]+$aClientSize[2]/2)-(1000/2), ($aClientSize[1]+$aClientSize[3]/2)-(1000/2), 1000 ,1000)					
				EndSwitch
        Case $List2  
            local $index=0
            $List2Sel=_GUICtrlListBox_GetCurSel($List2)
			Local $aClientSize = WinGetPos ($hRedBox);WinGetClientSize($hRedBox)
				Switch $List2Sel
					Case 0
						WinMove($hRedBox, "", 101, 101, $aClientSize[2] ,$aClientSize[2] )
					Case 1
						WinMove($hRedBox, "", 101, 101, $aClientSize[2]*1.5,$aClientSize[3] )
					Case 2
						WinMove($hRedBox, "", 101, 101, $aClientSize[2]*2 ,$aClientSize[3] )
					Case 3
						WinMove($hRedBox, "", 101, 101, $aClientSize[2]*3 ,$aClientSize[3])
					Case 4
						WinMove($hRedBox, "", 101, 101, $aClientSize[2]*4 ,$aClientSize[3])
					Case 5
						WinMove($hRedBox, "", 101, 101, $aClientSize[2]*5 ,$aClientSize[3])
					Case 6
						WinMove($hRedBox, "", 101, 101, $aClientSize[2]*6 ,$aClientSize[3])			
					Case 7
						WinMove($hRedBox, "", 101, 101, $aClientSize[2]*7 ,$aClientSize[3])
					Case 8
						WinMove($hRedBox, "", 101, 101, $aClientSize[2]*8 ,$aClientSize[3])					
				EndSwitch
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
	EndIf
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