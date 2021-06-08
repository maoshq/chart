#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <MsgBoxConstants.au3>
#include <Date.au3>

HotKeySet("{ESC}","Terminate")

while(1)

    Local $Excelobj = _Excel_Open(Default, Default, Default, Default, True)
    If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookOpen Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
    
    $Num = 1
    $open=FileOpen("log.txt",9)
    
    $Excelobj.Visible=1
    $Excelobj.WorkBooks.Add;
    $Excelobj.WindowState= -4137
    
    $WBName1 = $Excelobj.ActiveWorkbook.Name 
    Sleep(1000)
    _Excel_BookNew($Excelobj)
    $WBName2 = $Excelobj.ActiveWorkbook.Name 
    $Excelobj.Workbooks($WBName1).Activate 
    
    For $i = 0 To 5 Step +1
        
    
    Sleep(2000)
    Local $aPos = MouseGetPos()
    FileWrite($open,_NowTime()&"-----------  "&$WBName1&" : ("& $aPos[0] & ", " & $aPos[1]& ") --> ")
    
    MouseClick("left",300,550)
    Sleep(1000)
    FileWrite($open,"("& MouseGetPos()[0] & ", " & MouseGetPos()[1]&")  ")
    
    
    Sleep(2000)
    $Excelobj.Workbooks($WbName2).Activate
    Local $aPos = MouseGetPos()
    FileWrite($open,"  "&$WBName2&" : ("& $aPos[0] & ", " & $aPos[1]& ") --> ")
    MouseClick("left",550,300)
    Sleep(1000)
    FileWrite($open,"("& MouseGetPos()[0] & ", " & MouseGetPos()[1]&")  ")
    
    
    Sleep(3000)
    $Excelobj.Workbooks($WBName1).Activate 
    $WBName = $Excelobj.ActiveWorkbook.Name 
    
    FileWrite($open,"   WindowSwitch --> "&$WBName& @crlf)
    
    Next
    Sleep(2000)
    
    FileClose($open)
    $Excelobj.ActiveWorkBook.Saved = 1
    $Excelobj.Quit;
WEnd

Func Terminate()
    MsgBox("4096","","Process exit")
    Exit;
EndFunc