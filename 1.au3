#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <Word.au3>
#include <File.au3>

Local $startDay = @MDAY ,$startHour = @HOUR
_FileWriteLog(@ScriptDir&"/Log/log"&@MON&"-"&$startDay&""&$startHour&".txt","--------script start---------")
Func _Run_Edge()
    For $i = 0 To 10 Step +1
        OperEdge1()
        Sleep(3000)
    Next
    OperEdge()
    Sleep(3000)
EndFunc

Func _Run_PowerPoint()

    For $i = 0 To 5 Step +1
        $PPobj=ObjCreate("PowerPoint.Application")
        $PPobj.Presentations.Add
        $PPobj.Activate;
        $PPobj.WindowState = 3
        Sleep(3000)
        If WinExists("[CLASS:NUIDialog]") Then WinClose("[CLASS:NUIDialog]")
        Sleep(4000)
        _Pp1_AddTitle()
        Sleep(2000)
        _Pp1_AddNode(0)
        Sleep(2000)
        _Pp1_AddNode(1)
        Sleep(2000)
        _Pp1_AddNode(2)
        Sleep(2000)
        _Pp1_Play(4)
        Sleep(2000)
        $PPobj.quit
        Sleep(2000)
    Next
EndFunc

Func _RunExcel()
    For $i = 0 To 20 Step +1
        _Ex1_EditChart1()
        Sleep(3000)
        _Ex1_OperColumn(Random(10,20,1))
        Sleep(3000)
        _Ex1_OperateCell()
        Sleep(3000)
    Next

    _Ex1_OperColumn(Random(10,20,1),1)

    _Ex1_OperateCell1()
EndFunc

Func _RunWord()
    For $i = 0 To 20 Step +1
        _Wd1_Oper1()
        Sleep(3000)
    Next
    _Wd1_Oper()
    Sleep(2000)
EndFunc

Func Oper_Outlook()
    For $i = 0 To 10 Step +1
        Local $Olkobj=ObjCreate("Outlook.Application")
        $Olkobj.Explorers.Add($Olkobj.GetNamespace("MAPI").GetDefaultFolder("6"),"0").Display ;display inbox explorer
        $Olkobj.ActiveWindow.WindowState =0
        Sleep(5000)
        If WinExists("[CLASS:NUIDialog]") Then WinClose("[CLASS:NUIDialog]")
        Sleep(1000)
        $Olkobj.CreateItem("0").Display;          open newMail window
        Sleep(2000)
        Send("77829393612e@qq.com"&"{TAB 2}")
        Sleep(2000)
        Send(Random(100,1000,1)&"{TAB}")
        Sleep(2000)
        Send("Test:"&Random(1000,10000,1))
        Sleep(2000)
        Send("!{F4}")
        Sleep(2000)
        Send("{RIGHT}{ENTER}")
        Sleep(3000)
        $Olkobj.Quit
    Next
EndFunc
Func _Ex1_LocateCell($locate)
    Switch (Random(0,1,1))
        case 0
            Sleep(2000)
            Send("{F5}")
            Send($locate)
            sleep(2000)
            Send("{ENTER}")
        case 1
            Sleep(2000)
            _Ex1_LocateCell2($locate)
    EndSwitch
EndFunc

Func _Ex1_LocateCell2($locate)
    Sleep(1000)
    ControlClick("","","Edit1")
    Send($locate&"{Enter}")
EndFunc


Func _Ex1_EditChart1()
    $time1 = 1
    Local $Excelobj=ObjCreate("Excel.Application")
    Sleep(2000)
    $Excelobj.Visible=1

    $Excelobj.WorkBooks.Add;
    $Excelobj.WindowState= -4137
    Sleep(3000)
    If WinExists("[CLASS:NUIDialog]") Then WinClose("[CLASS:NUIDialog]")

    Sleep(1000)
   ;$Excelobj.ActiveWindow
    $le=Random(65,83,1)
    $x=Chr($le)
    $l=Random(1,30,1)
    _Ex1_LocateCell2($x&$l)
    For $i = 1 To 5 Step +1
        ;Sleep(2000)
        For $y = 0 To 4 Step +1
        Send(Random(0,100,1)&"{RIGHT}")
        ;Sleep(2000)
        Next
    $l+=1
    _Ex1_LocateCell2($x&$l)
    Next
    Local $ChartY = Random(1,7,1)
    _Ex1_LocateCell2($x&$l-5&":"&Chr($le+4)&$l-1)
    Send("{ALT}HT")
    Sleep(1000)
    Send("{DOWN " & $ChartY & "}"&"{RIGHT " & $ChartY & "}"&"{ENTER 2}")
    _Ex1_LocateCell2($x&$l-4&":"&Chr($le+4)&$l)
    Send("{ALT}MUA");                               求平均值
    _Ex1_LocateCell2($x&$l-4&":"&Chr($le+4)&$l-4)
    Send("{ALT}MUS");                               求和

    _Ex1_LocateCell2(Chr($le+5)&$l-4)
    Send("=sum("&$x&$l-4&":"&Chr($le+4)&$l-3&")"&"{ENTER}")  ;
    _Ex1_LocateCell2(Chr($le+5)&$l-4)
    Send("=sum("&$x&$l-2&":"&Chr($le+4)&$l-2&")"&"{ENTER}")

    Send("=sum("&$x&$l-1&":"&Chr($le+4)&$l-1&")"&"{ENTER}")
    ;Sleep(2000)
    Send("=sum("&$x&$l&":"&Chr($le+4)&$l&")"&"{ENTER}")

    _Ex1_LocateCell2($x&$l-2&":"&Chr($le+4)&$l-1)
    Send("{ALT}H1"&"^I")
    Sleep(2000)
    Send("+{F10}N")
    
    Sleep(4000)
    _FileWriteLog(@ScriptDir&"/Log/log"&@MON&"-"&@MDAY&"  "&@HOUR&".txt","")
    Sleep(1000)
    $Excelobj.ActiveWorkBook.Saved = 1
    $Excelobj.Quit;
EndFunc

Func _Ex1_ClickInSheet()
    WinActivate("[CLASS:XLMAIN]")
    $arr=ControlGetPos("[CLASS:XLMAIN]","","EXCEL71")
    $ayy1=WinGetPos("[CLASS:XLMAIN]")
    MouseClick("left",Random($arr[0]+$ayy1[0]+$arr[2]/10,$arr[0]+$ayy1[0]+$arr[2]*0.9),Random($arr[1]+$ayy1[1]+$arr[3]/10,$arr[1]+$ayy1[1]+$arr[3]*0.9))
EndFunc

Func _Ex1_OperSinCell()
    Local $Excelobj=ObjCreate("Excel.Application")
    Sleep(2000)
    $Excelobj.Visible=1

    $Excelobj.WorkBooks.Add;
    $Excelobj.WindowState= -4137
    Sleep(3000)
    If WinExists("[CLASS:NUIDialog]") Then WinClose("[CLASS:NUIDialog]")
    Sleep(1000)

    $RandomX=Random(65,84,1)

    $Letter=Chr($RandomX)
    $RandomY=Random(1,30,1);
    _Ex1_LocateCell2($Letter&$RandomY)

EndFunc
Func _Ex1_OperColumn($count,$isQuit=Default)
    Local $Excelobj=ObjCreate("Excel.Application")
    If $isQuit =Default Then
        $Excelobj.Visible=1
    
        $Excelobj.WorkBooks.Add;
        $Excelobj.WindowState= -4137
        Sleep(3000)
        If WinExists("[CLASS:NUIDialog]") Then WinClose("[CLASS:NUIDialog]")
        Sleep(3000)
    
        $RandomX=Random(65,84,1)
        $Letter=Chr($RandomX)
        $RandomY=Random(1,30,1);
    
        _Ex1_LocateCell2($Letter&$RandomY)
    
        For $i = 0 To $count Step +1
            ;Sleep(2000)
            Send(Random(0,1000,1)&"{DOWN}")
        Next
        Sleep(2000)
    
        Switch (Random(0,6,1))
            Case 0
                $sum=Chr(83)&Chr(85)&Chr(77);
                Send("="&$sum&"("&$Letter&$RandomY&":"&$Letter&$RandomY+$count&")"&"{ENTER}")
                ;Send("+{F3}")  ;调用函数  =SUM(F10,F17)
                Send("{UP}{ALT}H1")
            Case 1
                Send("=average("&$Letter&$RandomY&":"&$Letter&$RandomY+$count&")"&"{ENTER}")
                sleep(2000)
                Send("{UP}{ALT}H2")
            Case 2
                Send("=count("&$Letter&$RandomY&":"&$Letter&$RandomY+$count&")"&"{ENTER}")
                sleep(2000)
                Send("{UP}{ALT}H2")
            Case 3
                Send("=MAX("&$Letter&$RandomY&":"&$Letter&$RandomY+$count&")"&"{ENTER}")
                sleep(2000)
                Send("{UP}{ALT}H1")
            Case 4
                _Ex1_LocateCell2($Letter&$RandomY&":"&$Letter&$RandomY+$count)
                sleep(3000)
                Send("+{F10}OS")
            Case 5
                _Ex1_LocateCell2($Letter&$RandomY&":"&$Letter&$RandomY+$count)
                Send("^x")
                Sleep(1000)
                _Ex1_ClickInSheet()
                Sleep(2000)
                Send("^v")
            Case 6
                _Ex1_LocateCell2($Letter&$RandomY&":"&$Letter&$RandomY+$count)
                Send("{ALT}HT{TAB}{ENTER 2}")   ;biaogeyangshi
        EndSwitch
        Sleep(4000)
        $Excelobj.ActiveWorkBook.Saved = 1
        $Excelobj.Quit;
    EndIf

    If $isQuit =1 Then
        $Excelobj.Visible=1
    
        $Excelobj.WorkBooks.Add;
        $Excelobj.WindowState= -4137
        Sleep(3000)
        If WinExists("[CLASS:NUIDialog]") Then WinClose("[CLASS:NUIDialog]")
        Sleep(1000)

        For $z = 0 To 15 Step +1
        $RandomX=Random(65,84,1)
        $Letter=Chr($RandomX)
        $RandomY=Random(1,30,1);
    
        _Ex1_LocateCell2($Letter&$RandomY)
    
        For $i = 0 To $count Step +1
            ;Sleep(2000)
            Send(Random(0,1000,1)&"{DOWN}")
        Next
        Sleep(2000)
    
        Switch (Random(0,6,1))
            Case 0
                $sum=Chr(83)&Chr(85)&Chr(77);
                Send("="&$sum&"("&$Letter&$RandomY&":"&$Letter&$RandomY+$count&")"&"{ENTER}")
                ;Send("+{F3}")  ;调用函数  =SUM(F10,F17)

            Case 1
                Send("=average("&$Letter&$RandomY&":"&$Letter&$RandomY+$count&")"&"{ENTER}")
                sleep(2000)

            Case 2
                Send("=count("&$Letter&$RandomY&":"&$Letter&$RandomY+$count&")"&"{ENTER}")
                sleep(2000)

            Case 3
                Send("=MAX("&$Letter&$RandomY&":"&$Letter&$RandomY+$count&")"&"{ENTER}")
                sleep(2000)

            Case 4
                _Ex1_LocateCell2($Letter&$RandomY&":"&$Letter&$RandomY+$count)
                sleep(3000)
                Send("+{F10}OS")
            Case 5
                _Ex1_LocateCell2($Letter&$RandomY&":"&$Letter&$RandomY+$count)
                Send("^x")
                Sleep(1000)
                _Ex1_ClickInSheet()
                Sleep(2000)
                Send("^v")
            Case 6
                _Ex1_LocateCell2($Letter&$RandomY&":"&$Letter&$RandomY+$count)
                Send("{ALT}HT{TAB}{ENTER 2}")   ;biaogeyangshi
        EndSwitch
        Sleep(2000)
        Send("^a"&"+{F10}N")
        Next
        Sleep(4000)
        $Excelobj.ActiveWorkBook.Saved = 1
        $Excelobj.Quit;
    EndIf
EndFunc

Func _Ex1_SaveAndDel()
    Local $Excelobj=ObjCreate("Excel.Application")
    $Excelobj.Visible=1                                     ; 显示 Excel 自己
    $Excelobj.WorkBooks.Add;
    $Excelobj.ActiveWorkBook.ActiveSheet.Cells(5,5).Value="test" ;
    send("{ALT}FS")
    Sleep(2000)
    send("{ENTER 3}")
    Sleep(2000)
    Send("{BACKSPACE}")
    send(@DesktopDir&"\1.xlsx")
    Sleep(2000)
    Send("{ENTER 2}")
    Sleep(2000)
    $Excelobj.Quit;
    Sleep(4000)
    FileDelete(@DesktopDir&"\1.xlsx")
EndFunc


Func _Ex1_OperateCell()
    Local $Excelobj=ObjCreate("Excel.Application")
    Sleep(2000)
    $Excelobj.Visible=1

    $Excelobj.WorkBooks.Add;
    $Excelobj.WindowState= -4137
    Sleep(3000)
    If WinExists("[CLASS:NUIDialog]") Then WinClose("[CLASS:NUIDialog]")
    Sleep(1000)

    _Ex1_ClickInSheet()
    Send(Random(0,10000,1)&"{ENTER}{UP}")
    Switch (Random(0,2,1))
        case 0
            Send("^x")
            Sleep(2000)
            _Ex1_ClickInSheet()
            Sleep(2000)
            Send("^v")
        Case 1
            Send("^b")
            Sleep(2000)
            Send("^i")
        Case 2
            Send("{ALT}HBT")
            Sleep(2000)
            Send("+{F10}D{ENTER}")
    EndSwitch
    Sleep(2000)
    $Excelobj.ActiveWorkBook.Saved = 1
    $Excelobj.Quit;
EndFunc

Func _Ex1_OperateCell1()
    Local $Excelobj=ObjCreate("Excel.Application")
    Sleep(2000)
    $Excelobj.Visible=1

    $Excelobj.WorkBooks.Add;
    $Excelobj.WindowState= -4137
    Sleep(3000)
    If WinExists("[CLASS:NUIDialog]") Then WinClose("[CLASS:NUIDialog]")
    Sleep(1000)

    For $i = 0 To 20 Step +1
        $Excelobj.Visible=1
        _Ex1_ClickInSheet()
        Send(Random(0,10000,1)&"{ENTER}{UP}")
        Switch (Random(0,2,1))
            case 0
                Send("^x")
                Sleep(2000)
                _Ex1_ClickInSheet()
                Sleep(2000)
                Send("^v")
                sleep(2000)
                Send("+{F10}N")
                Sleep(2000)
            Case 1
                Send("^b")
                Sleep(2000)
                Send("^i")
                Sleep(2000)
                Send("+{F10}N")
                Sleep(2000)
            Case 2
                Send("{ALT}HBT")
                Sleep(2000)
                Send("+{F10}D{ENTER}")
                Sleep(2000)
        EndSwitch
    Next
    $Excelobj.ActiveWorkBook.Saved = 1
    $Excelobj.Quit;
EndFunc


;              Word
Func _Wd1_Oper($IsQuit = Default)
    Local $ChartX = Random(1,15,1),$ChartY = Random(1,7,1),$ChartZ=Random(1,10,1)
    Local $oWord=ObjCreate("Word.Application")
	$oWord.Visible=True
    $oWord.Documents.Add()
    $oWord.WindowState = 1
    $oWord.Activate
    Sleep(3000)
    If WinExists("[CLASS:NUIDialog]") Then WinClose("[CLASS:NUIDialog]")
    Sleep(1000)

    Send("{ALT}HL{RIGHT 2}{ENTER}")
    Sleep(2000)
    Send("{ALT}HAC"&"Title One"&"{ENTER}")
    Sleep(2000)
    Send("{ALT}HFS"&Random(15,26,1)&"{ENTER}")
    Sleep(2000)
    Send("{ALT}HFC"&"{DOWN " & $ChartY & "}"&"{RIGHT " & $ChartZ & "}"&"{ENTER}")
    Sleep(1000)
    Send("^b"&"^i")
    Sleep(1000)
    For $i = 0 To 3 Step +1
        For $y = 0 To Random(10,20,1) Step +1
            Send(Chr(Random(65,90,1)))
        Next
        Sleep(2000)
        Send("{ENTER}")
    Next

    For $z = 0 To 20 Step +1
        Switch (Random(0,3,1))
            Case 0
                Send("+{F10}H"&"www.baidu.com"&"{ENTER}")
                Sleep(2000)
                Send("{ALT}HSLA{BACKSPACE}")
                Sleep(1000)
            Case 1
                For $i = 0 To 3 Step +1
                    For $y = 0 To Random(10,20,1) Step +1
                        Send(Chr(Random(65,90,1)))
                    Next
                    Send("{ENTER}")
                Next
                Sleep(2000)
                Send("{ALT}HSLA"&"^c"&"{BACKSPACE}")
                Sleep(2000)
                Send("^v")
                Sleep(2000)
                Send("^a{BACKSPACE}")   ;  delete
                Sleep(1000)
            Case 2
                For $i = 0 To 2 Step +1
                    For $y = 0 To Random(10,20,1) Step +1
                        Send(Chr(Random(65,90,1)))
                    Next
                    Send("{ENTER}")
                Next
                Sleep(2000)
                Send("{ALT}HSLA"&"{ALT}HI{ENTER}")
                Sleep(2000)
                Send("^a{BACKSPACE}")
                Sleep(1000)
            Case 3
                Send("{ALT}HSLA{BACKSPACE}")   ;  delete
                Sleep(2000)
                Send("{ALT}NT{RIGHT 4}{DOWN 4}{ENTER}")
                Sleep(2000)
                Send("{ALT}JTS"&"{DOWN " & $ChartY & "}"&"{RIGHT " & $ChartX & "}"&"{ENTER}")
                Sleep(2000)
                For $i = 0 To Random(6,20,1) Step +1
                    Send(Random(0,10,1)&"{RIGHT}")
                Next
                Send("^a{BACKSPACE}")
                Sleep(1000)
        EndSwitch
    Next

    Send("{ALT}HSLA{BACKSPACE}")   ;  delete
    Sleep(2000)

    If $IsQuit =Default Then
        $oWord.Documents.Close(0)
        $oWord.Quit
    EndIf
EndFunc

Func _Wd1_Oper1()

    Local $ChartX = Random(1,15,1),$ChartY = Random(1,7,1),$ChartZ=Random(1,10,1)
    Local $oWord=ObjCreate("Word.Application")
	$oWord.Visible=True
    $oWord.Documents.Add()
    $oWord.WindowState = 1
    $oWord.Activate
    Sleep(3000)
    If WinExists("[CLASS:NUIDialog]") Then WinClose("[CLASS:NUIDialog]")
    Sleep(1000)

    Send("{ALT}HL{RIGHT 2}{ENTER}")
    Sleep(2000)
    Send("{ALT}HAC"&"Title One"&"{ENTER}")
    Sleep(2000)
    Send("{ALT}HFS"&Random(15,26,1)&"{ENTER}")
    Sleep(2000)
    Send("{ALT}HFC"&"{DOWN " & $ChartY & "}"&"{RIGHT " & $ChartZ & "}"&"{ENTER}")
    Sleep(1000)
    Send("^b"&"^i")
    Sleep(1000)
    For $t = 0 To 3 Step +1
        For $y = 0 To Random(10,20,1) Step +1
            Send(Chr(Random(65,90,1)))
        Next
        Send("{ENTER}")
    Next
    Switch (Random(0,3,1))
        Case 0
            Send("+{F10}H"&"www.baidu.com"&"{ENTER}")
            Sleep(2000)
            Send("^a{BACKSPACE}")
        Case 1
            For $i = 0 To 3 Step +1
                For $y = 0 To Random(10,20,1) Step +1
                    Send(Chr(Random(65,90,1)))
                Next
                Send("{ENTER}")
            Next
            Sleep(2000)
            Send("{ALT}HSLA"&"^c"&"{BACKSPACE}")
            Sleep(2000)
            Send("^v")
            Sleep(2000)
            Send("^a{BACKSPACE}")   ;  delete
            Sleep(2000)
        Case 2
            Send("{ALT}HSLA{BACKSPACE}")
            Send("+{F10}MM{ENTER}")
            Sleep(2000)
            Send(Chr(Random(65,90,1))&Random(65,90,1))
            Sleep(2000)

        #comments-start
                    Case 3
            Send(Random(0,9,1)&"{ALT}HSLA"&"{ALT}HI{ENTER}")
            Sleep(2000)
            Send("{ALT}HSLA{BACKSPACE}")
        #comments-end

        Case 3
            Send("{ALT}HSLA{BACKSPACE}")   ;  delete
            Sleep(2000)
            Send("{ALT}NT{RIGHT 4}{DOWN 4}{ENTER}")
            Sleep(2000)
            Send("{ALT}JTS"&"{DOWN " & $ChartY & "}"&"{RIGHT " & $ChartX & "}"&"{ENTER}")
            Sleep(2000)
            For $i = 0 To Random(6,20,1) Step +1
                Send(Random(0,10,1)&"{RIGHT}")
            Next
            Send("^a{BACKSPACE}")
    EndSwitch
    Send("^a{BACKSPACE}")   ;  delete
    Sleep(2000)

    $oWord.Documents.Close(0)
    $oWord.Quit
EndFunc

;              Edge
Func OperEdge()
    Run(@ComSpec & " /c Explorer shell:AppsFolder\Microsoft.MicrosoftEdge_8wekyb3d8bbwe!App")
    ;Run("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
    Sleep(2000)
    For $i = 0 To 10 Step +1
        Send("www.baidu.com")
        Sleep(1000)
        Send("{ENTER}")
        Sleep(3000)
        Send(Random(0,1000,1)&"{ENTER}")
        Sleep(3000)
        Send("^t")
    Next
    Send("www.baidu.com")
    Sleep(1000)
    Send("{ENTER}")
    Sleep(3000)
    Send(Random(0,1000,1)&"{ENTER}")
    Sleep(2000)
    Send("!{F4}")
EndFunc

Func OperEdge1()
    Run(@ComSpec & " /c Explorer shell:AppsFolder\Microsoft.MicrosoftEdge_8wekyb3d8bbwe!App")
    ;Run("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
    Sleep(3000)
    Send("www.baidu.com")
    Sleep(1000)
    Send("{ENTER}")
    Sleep(3000)
    Send(Random(0,1000,1)&"{ENTER}")
    Sleep(3000)
    Send("!fs")
    Sleep(3000)
    Send("!{F4}")
EndFunc

;              PowerPoint
Func _Pp1_AddTitle()
    Global $ColorX = Random(1,10,1),$ColorY = Random(1,5,1) ;Send("+{TAB " & $n & "}")
    Sleep(2000)
    Send("^m") ;Add Title 1
    Send(Random(10,99,1))
    Sleep(2000)
    Send("^+{LEFT}")
    Sleep(2000)
    Send("^b"&"+{F10}X{TAB}")
    Sleep(2000)
    Send(Random(10000,99999,1))
    Sleep(2000)
    Send("^+{LEFT}")
    Sleep(2000)
    Send("^u"&"{ALT}HFC"&"{RIGHT " & $ColorX & "}"&"{DOWN " & $ColorY & "}"&"{ENTER}") ;choose Random Color
    Sleep(2000)
    _Pp1_SetAnmi(0)
EndFunc

Func _Pp1_AddNode($serial)
    Global $ChartX = Random(1,15,1),$ChartY = Random(1,7,1)
    Switch ($serial)
        Case 0
            Sleep(2000)
            Send("^m")
            Sleep(2000)
            Send("Title:"&Random(0,9,1))
            Sleep("1000")
            Send("+{F10}X{TAB}")
            Sleep(1000)
            Send(Random(0,9,1)&"{ALT}NC{TAB}"&"{DOWN " & $ChartX & "}"&"{ENTER}")
            Sleep(5000)
            Send("!{F4}")
            Sleep(2000)
            _Pp1_SetAnmi(1)
            Sleep(2000)
            _Pp1_SetAnmi(0)
        Case 1
            Sleep(2000)
            Send("^m")
            Sleep(2000)
            Send("Title:"&Random(0,9,1))
            ;_Pp1_SetAnmi(1)
            Sleep("1000")
            Send("+{F10}X{TAB}")
            Sleep(1000)
            Send(Random(0,9,1)&"{ALT}NM{TAB}"&"{DOWN " & $ChartY & "}"&"{RIGHT " & $ChartX & "}"&"{ENTER}")
            Sleep(2000)
            _Pp1_SetAnmi(1)
            Sleep(2000)
            _Pp1_SetAnmi(0)
        Case 2
            Sleep(2000)
            Send("^m")
            Sleep(2000)
            Send("Title:"&Random(0,9,1))
            ;_Pp1_SetAnmi(1)
            Sleep("1000")
            Send("+{F10}X{TAB}")
            Sleep(1000)
            Send(Random(0,9,1)&"{ALT}NT{TAB}"&"{RIGHT 5}"&"{DOWN 3}"&"{ENTER}")
            Sleep(2000)
            Send("{ALT}JTA"&"{UP " & $ColorY & "}"&"{RIGHT " & $ColorY & "}"&"{ENTER}")
            Sleep(2000)
            For $i = 0 To 5 Step +1
                Send($i&"{TAB}")
                Sleep(1000)
            Next
            _Pp1_SetAnmi(1)
            Sleep(2000)
            _Pp1_SetAnmi(0)
            Sleep(2000)
    EndSwitch
EndFunc


Func _Pp1_Play($sum)
    Sleep(2000)
    Send("{F5}")
    For $i = 0 To $sum+1 Step +1
        Sleep(4000)
        Send("{ENTER}")
    Next
    Sleep(2000)
EndFunc


Func _Pp1_SetAnmi($Model)
    Local $AnmiX = Random(1,15,1),$AnmiY = Random(1,5,1)

    Switch ($Model)
        Case 0
            Sleep(1000)
            Send("{ALT}KT")
            Sleep(2000)
            Send("{RIGHT " & $AnmiX & "}"&"{DOWN " & $AnmiY & "}"&"{ENTER}");
            Sleep(3000)
        Case 1
            Sleep(1000)
            Send("{ALT}AS")
            Sleep(2000)
            Send("{RIGHT " & $AnmiX & "}"&"{DOWN " & $AnmiY & "}"&"{ENTER}");
            Sleep(3000)
    EndSwitch
EndFunc
