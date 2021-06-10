#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <MsgBoxConstants.au3>
#include <Date.au3>

HotKeySet("{ESC}","Terminate")

Func _Ex1_ClickInSheet()
    WinActivate("[CLASS:XLMAIN]")
    $arr=ControlGetPos("[CLASS:XLMAIN]","","EXCEL71")
    $ayy1=WinGetPos("[CLASS:XLMAIN]")
    MouseClick("left",Random($arr[0]+$ayy1[0]+$arr[2]/10,$arr[0]+$ayy1[0]+$arr[2]*0.9),Random($arr[1]+$ayy1[1]+$arr[3]/10,$arr[1]+$ayy1[1]+$arr[3]*0.9))
EndFunc

Local $Excelobj = _Excel_Open(Default, Default, Default, Default, True)
$Num = 1
$open=FileOpen("log.txt",9)

$Excelobj.Visible=1
$Excelobj.WorkBooks.Add;
$Excelobj.WindowState= -4137
$WBName1 = $Excelobj.ActiveWorkbook.Name 
Sleep(5000)
Send("!{F4}");
Sleep(1000)
$Excelobj.Workbooks($WBName1).Activate 

While(1)
    
    For $i = 0 To 1000 Step +1

        For $i = 0 To Random(5,10) Step +1
            _Ex1_ClickInSheet()
            Send("{A 10}")
        Next
        MouseMove(150,300,3)
        Local $aPos1 = MouseGetPos()
        MouseMove(400,400,3)
        Local $aPos2 = MouseGetPos()
        
        MouseMove(150,300,3)
        Local $aPos3 = MouseGetPos()
        MouseMove(400,400,3)
        Local $aPos4 = MouseGetPos()

        MouseMove(150,300,3)
        Local $aPos5 = MouseGetPos()
        MouseMove(400,400,3)
        Local $aPos6 = MouseGetPos()

        MouseMove(150,300,3)
        Local $aPos7 = MouseGetPos()
        MouseMove(400,400,3)
        Local $aPos8 = MouseGetPos()

        FileWrite($open,_NowTime()&"-----------  ("& $aPos1[0] & ", " & $aPos1[1]& ") --> ("&$aPos2[0] & ", " & $aPos2[1]& ") --> ("&$aPos3[0] & ", " & $aPos3[1]& ") --> ("&$aPos4[0] & ", " & $aPos4[1]& ") --> ("&$aPos5[0] & ", " & $aPos5[1]& ") --> ("&$aPos6[0] & ", " & $aPos6[1]& ") --> ("&$aPos7[0] & ", " & $aPos7[1]& ") --> ("&$aPos8[0] & ", " & $aPos8[1]& ") " & @crlf)
        
    Next


WEnd

Func Terminate()
    MsgBox("4096","","Process exit")
    Exit;
EndFunc